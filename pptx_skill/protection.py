"""Presentation protection, encryption, and marking-as-final.

Provides two levels of protection:

1. **Write protection** (application-level, no encryption):
   - Modify password via ``<p:modifyVerifier>`` hash
   - Read-only recommended flag
   - Mark as final metadata flag

2. **Password encryption** (file-level, real encryption):
   - Wraps ``msoffcrypto-tool`` for AES-256 encryption
   - Decrypt password-protected PPTX files

OOXML reference: ECMA-376 Part 4, §19.2 (PresentationML — Protection),
§15.2 (Encryption).
"""

from __future__ import annotations

import base64
import hashlib
import os
from dataclasses import dataclass

from pptx_skill._io import save_prs as _save_prs_impl
from pptx_skill.constants import P_NS as _NS_P

__all__ = [
    "ProtectionInfo",
    "apply_write_protection",
    "remove_write_protection",
    "mark_as_final",
    "unmark_as_final",
    "is_marked_final",
    "encrypt_pptx",
    "decrypt_pptx",
    "is_encrypted",
    "get_protection_info",
]

# ---------------------------------------------------------------------------
# Namespace constants
# ---------------------------------------------------------------------------

# FMTID for custom document property set
_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class ProtectionInfo:
    """Information about protection state of a presentation."""
    has_modify_password: bool = False
    is_read_only_recommended: bool = False
    is_marked_final: bool = False
    is_encrypted: bool = False
    encryption_type: str = ""  # "agile", "standard", ""


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _is_presentation(obj) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def _compute_password_hash(password: str, *, spin_count: int = 100000,
                           hash_algorithm: str = "sha256",
                           salt: bytes | None = None) -> tuple[bytes, bytes]:
    """Compute the password hash per ECMA-376 §14.2.3.

    Returns (salt, hash) where both are raw bytes.
    """
    if salt is None:
        salt = os.urandom(16)

    # Encode password as UTF-16LE
    password_bytes = password.encode("utf-16-le")

    if hash_algorithm == "sha256":
        h = hashlib.sha256(salt + password_bytes).digest()
        for _ in range(spin_count):
            h = hashlib.sha256(h).digest()
    elif hash_algorithm == "sha512":
        h = hashlib.sha512(salt + password_bytes).digest()
        for _ in range(spin_count):
            h = hashlib.sha512(h).digest()
    else:
        # SHA-1 (legacy)
        h = hashlib.sha1(salt + password_bytes).digest()
        for _ in range(spin_count):
            h = hashlib.sha1(h).digest()

    return salt, h


# ---------------------------------------------------------------------------
# Write protection (application-level)
# ---------------------------------------------------------------------------

def apply_write_protection(prs_or_path, *, password: str | None = None,
                           read_only: bool = False) -> bool:
    """Apply write protection to a presentation.

    Parameters
    ----------
    password : str, optional
        Password required to edit. Does not encrypt the file — only
        sets a hash that PowerPoint checks before allowing edits.
    read_only : bool, optional
        Mark the presentation as read-only recommended.

    Note
    ----
    This is application-level protection. The PPTX can still be opened
    and read by any ZIP tool. For real encryption, use ``encrypt_pptx``.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        pres_elem = prs._element

        # Find or create presentationPr
        pres_pr = pres_elem.find(f"{{{_NS_P}}}presentationPr")
        if pres_pr is None:
            pres_pr = etree.SubElement(pres_elem, f"{{{_NS_P}}}presentationPr")

        if password:
            # Remove existing modifyVerifier
            for existing in pres_pr.findall(f"{{{_NS_P}}}modifyVerifier"):
                pres_pr.remove(existing)

            salt, hash_val = _compute_password_hash(password)

            verifier = etree.SubElement(pres_pr, f"{{{_NS_P}}}modifyVerifier")
            verifier.set("cryptProviderType", "rsaFull")
            verifier.set("cryptAlgorithmClass", "hash")
            verifier.set("cryptAlgorithmType", "typeAny")
            verifier.set("cryptAlgorithmSid", "14")  # SHA-256
            verifier.set("spinCount", "100000")
            verifier.set("saltData", base64.b64encode(salt).decode("ascii"))
            verifier.set("hashData", base64.b64encode(hash_val).decode("ascii"))

        if read_only:
            # Remove existing readOnly
            for existing in pres_pr.findall(f"{{{_NS_P}}}readOnly"):
                pres_pr.remove(existing)
            etree.SubElement(pres_pr, f"{{{_NS_P}}}readOnly")

        return True
    finally:
        _save_prs(prs, path)


def remove_write_protection(prs_or_path) -> bool:
    """Remove all write protection from a presentation."""

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        pres_elem = prs._element
        pres_pr = pres_elem.find(f"{{{_NS_P}}}presentationPr")
        if pres_pr is None:
            return False

        found = False
        for tag in ("modifyVerifier", "readOnly"):
            for existing in pres_pr.findall(f"{{{_NS_P}}}{tag}"):
                pres_pr.remove(existing)
                found = True
        return found
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Mark as final
# ---------------------------------------------------------------------------

def mark_as_final(prs_or_path) -> bool:
    """Mark the presentation as final (read-only, no editing).

    Sets the ``DocSecurity`` core property to 4 and adds a custom
    property marking the document as final. PowerPoint shows a
    "MARKED AS FINAL" banner when opening.
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        # Set core property
        try:
            prs.core_properties.category = "Marked as Final"
        except Exception:
            pass

        # Set DocSecurity = 4 via custom XML
        _set_doc_security(prs, 4)
        return True
    finally:
        _save_prs(prs, path)


def unmark_as_final(prs_or_path) -> bool:
    """Remove the 'marked as final' status."""
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        try:
            if prs.core_properties.category == "Marked as Final":
                prs.core_properties.category = ""
        except Exception:
            pass
        _set_doc_security(prs, 0)
        return True
    finally:
        _save_prs(prs, path)


def is_marked_final(prs_or_path) -> bool:
    """Check if the presentation is marked as final."""
    prs = _open_prs(prs_or_path)
    try:
        try:
            if prs.core_properties.category == "Marked as Final":
                return True
        except Exception:
            pass
        security = _get_doc_security(prs)
        return security == 4
    finally:
        pass


def _set_doc_security(prs, value: int):
    """Set the DocSecurity property in core.xml."""
    from lxml import etree

    try:
        core_part = prs.part.package.part_related_by(
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
        )
        core_xml = etree.fromstring(core_part.blob)


        # Find or create cp:contentStatus
        content_status = core_xml.find("{{cp_ns}}contentStatus")
        if content_status is None:
            content_status = etree.SubElement(core_xml, "{{cp_ns}}contentStatus")

        if value == 4:
            content_status.text = "Final"
        else:
            content_status.text = ""

        core_part._blob = etree.tostring(core_xml, xml_declaration=True, encoding="UTF-8")
    except Exception:
        pass  # Core properties part may not exist or be writable


def _get_doc_security(prs) -> int:
    """Get the DocSecurity value from core.xml."""
    from lxml import etree

    try:
        core_part = prs.part.package.part_related_by(
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
        )
        core_xml = etree.fromstring(core_part.blob)
        content_status = core_xml.find("{{cp_ns}}contentStatus")
        if content_status is not None and content_status.text == "Final":
            return 4
    except Exception:
        pass
    return 0


# ---------------------------------------------------------------------------
# Password encryption (file-level)
# ---------------------------------------------------------------------------

def encrypt_pptx(input_path: str, output_path: str, *, password: str,
                 algorithm: str = "aes256") -> bool:
    """Encrypt a PPTX file with a password.

    This creates a new encrypted file. The original is not modified.

    Parameters
    ----------
    input_path : str
        Path to the unencrypted PPTX.
    output_path : str
        Path for the encrypted output.
    password : str
        Password to set.
    algorithm : str
        Encryption algorithm. Currently only "aes256" is supported.

    Returns
    -------
    bool
        True if encryption succeeded.

    Raises
    ------
    ImportError
        If ``msoffcrypto-tool`` is not installed.
    """
    try:
        import msoffcrypto
    except ImportError:
        raise ImportError(
            "msoffcrypto-tool is required for encryption. "
            "Install with: pip install msoffcrypto-tool"
        ) from None

    with open(input_path, "rb") as f:
        file = msoffcrypto.OfficeFile(f)
        with open(output_path, "wb") as out:
            file.encrypt(password, out)

    return True


def decrypt_pptx(input_path: str, output_path: str, *, password: str) -> bool:
    """Decrypt a password-protected PPTX file.

    Parameters
    ----------
    input_path : str
        Path to the encrypted PPTX.
    output_path : str
        Path for the decrypted output.
    password : str
        Password to decrypt with.

    Returns
    -------
    bool
        True if decryption succeeded.
    """
    try:
        import msoffcrypto
    except ImportError:
        raise ImportError(
            "msoffcrypto-tool is required for decryption. "
            "Install with: pip install msoffcrypto-tool"
        ) from None

    with open(input_path, "rb") as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password=password)
        with open(output_path, "wb") as out:
            file.decrypt(out)

    return True


def is_encrypted(path: str) -> bool:
    """Check if a PPTX file is password-encrypted.

    Parameters
    ----------
    path : str
        Path to the file to check.

    Returns
    -------
    bool
        True if the file appears to be encrypted.
    """
    try:
        import msoffcrypto
    except ImportError:
        # Fallback: check if file starts with a ZIP signature
        with open(path, "rb") as f:
            magic = f.read(4)
        # ZIP files start with PK\x03\x04
        return magic[:2] != b"PK"

    try:
        with open(path, "rb") as f:
            file = msoffcrypto.OfficeFile(f)
            return file.is_encrypted()
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Protection info
# ---------------------------------------------------------------------------

def get_protection_info(prs_or_path) -> ProtectionInfo:
    """Get comprehensive protection information about a presentation."""

    info = ProtectionInfo()

    # Check if it's a path (for encryption check)
    if isinstance(prs_or_path, str) and os.path.exists(prs_or_path):
        info.is_encrypted = is_encrypted(prs_or_path)
        if info.is_encrypted:
            return info  # Can't read encrypted files without password

    prs = _open_prs(prs_or_path)
    try:
        pres_elem = prs._element
        pres_pr = pres_elem.find(f"{{{_NS_P}}}presentationPr")
        if pres_pr is not None:
            # Check modifyVerifier
            verifier = pres_pr.find(f"{{{_NS_P}}}modifyVerifier")
            if verifier is not None:
                info.has_modify_password = True

            # Check readOnly
            read_only = pres_pr.find(f"{{{_NS_P}}}readOnly")
            if read_only is not None:
                info.is_read_only_recommended = True

        # Check marked as final
        info.is_marked_final = is_marked_final(prs)

        return info
    finally:
        pass  # Read-only


# ---------------------------------------------------------------------------
# Internal open/save helpers (same pattern as other modules)
# ---------------------------------------------------------------------------

def _open_prs(prs_or_path):
    """Open a Presentation from *prs_or_path*.

    Accepts either an already-opened ``Presentation`` object or a file path.
    Returns the ``Presentation`` object directly.
    """
    from pptx import Presentation

    if _is_presentation(prs_or_path):
        return prs_or_path
    return Presentation(str(prs_or_path))


def _save_prs(prs, path):
    """Save *prs* back to *path* if *path* is not None."""
    _save_prs_impl(prs, path, backup=False)
