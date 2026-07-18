"""
pixabay_search.py - Pixabay 图片搜索与下载工具

提供 CLI 和 Python API 两种调用方式，用于从 Pixabay 搜索并下载
免费可商用的高质量图片，供 PPT 制作时作为素材使用。

Pixabay API 文档: https://pixabay.com/api/docs/
"""

import argparse
import os
import sys
import json
import time
import urllib.parse
import urllib.request
import urllib.error

# 默认 API Key（用户提供的示例 key）
DEFAULT_API_KEY = "56736282-f8f7357e3a0c8aca65350f06b"

API_ENDPOINT = "https://pixabay.com/api/"

# 合法的参数取值（用于校验）
VALID_IMAGE_TYPES = {"all", "photo", "illustration", "vector"}
VALID_ORIENTATIONS = {"all", "horizontal", "vertical"}
VALID_CATEGORIES = {
    "backgrounds", "fashion", "nature", "science", "education",
    "feelings", "health", "people", "religion", "places", "animals",
    "industry", "computer", "food", "sports", "transportation",
    "travel", "buildings", "business", "music"
}
VALID_LANGS = {
    "cs", "da", "de", "en", "es", "fr", "id", "it", "hu", "nl",
    "no", "pl", "pt", "ro", "sk", "fi", "sv", "tr", "vi", "th",
    "bg", "ru", "el", "ja", "ko", "zh"
}
VALID_ORDERS = {"popular", "latest"}
VALID_COLORS = {
    "grayscale", "transparent", "red", "orange", "yellow", "green",
    "turquoise", "blue", "lilac", "pink", "white", "gray", "black", "brown"
}


def search_images(
    query,
    api_key=DEFAULT_API_KEY,
    lang="zh",
    image_type="photo",
    orientation="horizontal",
    category=None,
    min_width=0,
    min_height=0,
    colors=None,
    editors_choice=None,
    safesearch=True,
    order="popular",
    page=1,
    per_page=20,
    count=None,
    timeout=15,
):
    """
    调用 Pixabay API 搜索图片，返回结果列表。

    参数:
        query: 搜索关键词（如 "artificial intelligence"）
        api_key: Pixabay API key
        lang: 搜索语言（默认 zh）
        image_type: 图片类型 all/photo/illustration/vector
        orientation: 方向 all/horizontal/vertical
        category: 分类（见 VALID_CATEGORIES）
        min_width: 最小宽度（像素）
        min_height: 最小高度（像素）
        colors: 颜色过滤（逗号分隔字符串或列表）
        editors_choice: True/False/None（None 表示不传）
        safesearch: 是否启用安全搜索
        order: popular/latest
        page: 页码
        per_page: 每页结果数（3-200）
        count: 仅返回前 N 条结果（如果指定）
        timeout: 请求超时秒数

    返回:
        list[dict]: 图片信息列表，每个 dict 包含:
            - id: 图片 ID
            - tags: 标签
            - previewURL: 预览图 URL
            - webformatURL: 中等尺寸图 URL（推荐下载）
            - largeImageURL: 大图 URL
            - imageWidth, imageHeight: 尺寸
            - views, downloads, likes, favorites: 统计数据
            - user: 作者
            - pageURL: Pixabay 页面链接
    """
    # 参数校验
    if image_type not in VALID_IMAGE_TYPES:
        raise ValueError(f"image_type 必须是 {VALID_IMAGE_TYPES} 之一")
    if orientation not in VALID_ORIENTATIONS:
        raise ValueError(f"orientation 必须是 {VALID_ORIENTATIONS} 之一")
    if category is not None and category not in VALID_CATEGORIES:
        raise ValueError(f"category 必须是 {VALID_CATEGORIES} 之一")
    if lang not in VALID_LANGS:
        raise ValueError(f"lang 必须是 {VALID_LANGS} 之一")
    if order not in VALID_ORDERS:
        raise ValueError(f"order 必须是 {VALID_ORDERS} 之一")
    if not (3 <= per_page <= 200):
        raise ValueError("per_page 必须在 3-200 之间")
    if colors is not None:
        if isinstance(colors, (list, tuple)):
            colors = ",".join(colors)
        color_list = [c.strip() for c in colors.split(",") if c.strip()]
        for c in color_list:
            if c not in VALID_COLORS:
                raise ValueError(f"颜色 '{c}' 无效，合法值: {VALID_COLORS}")

    # 构造请求参数
    params = {
        "key": api_key,
        "q": query,
        "lang": lang,
        "image_type": image_type,
        "orientation": orientation,
        "min_width": str(min_width),
        "min_height": str(min_height),
        "safesearch": "true" if safesearch else "false",
        "order": order,
        "page": str(page),
        "per_page": str(per_page),
    }
    if category:
        params["category"] = category
    if colors:
        params["colors"] = colors
    if editors_choice is not None:
        params["editors_choice"] = "true" if editors_choice else "false"

    url = API_ENDPOINT + "?" + urllib.parse.urlencode(params)

    # 发起请求
    req = urllib.request.Request(url, headers={"User-Agent": "pptx-skill/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Pixabay API 返回 HTTP {e.code}: {body}") from None
    except urllib.error.URLError as e:
        raise RuntimeError(f"无法连接 Pixabay API: {e.reason}") from None

    hits = data.get("hits", [])
    if count is not None:
        hits = hits[:count]
    return hits


def download_images(hits, output_dir="./images", size="webformat", timeout=30):
    """
    下载搜索结果中的图片到本地。

    参数:
        hits: search_images 返回的列表
        output_dir: 保存目录（不存在会自动创建）
        size: 下载哪种尺寸
            - "preview"    : 预览图（最小，约 150x150）
            - "webformat"  : 中等尺寸（推荐，约 640px，适合 PPT）
            - "large"      : 大图（原始尺寸，文件较大）
            - "fullHD"     : 全高清（部分图片提供）
        timeout: 下载超时秒数

    返回:
        list[str]: 成功下载的本地文件路径列表（顺序与 hits 对应，失败项跳过）
    """
    size_field = {
        "preview": "previewURL",
        "webformat": "webformatURL",
        "large": "largeImageURL",
        "fullHD": "fullHDURL",
    }.get(size, "webformatURL")

    os.makedirs(output_dir, exist_ok=True)
    downloaded = []

    for hit in hits:
        img_url = hit.get(size_field) or hit.get("webformatURL")
        if not img_url:
            continue

        img_id = hit.get("id", "unknown")
        # 从 URL 推断扩展名，默认 jpg
        ext = "jpg"
        for candidate in (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"):
            if candidate in img_url.lower():
                ext = candidate.lstrip(".")
                break
        filename = f"img_{img_id}.{ext}"
        filepath = os.path.join(output_dir, filename)

        # 已存在则跳过下载
        if os.path.exists(filepath) and os.path.getsize(filepath) > 0:
            downloaded.append(filepath)
            continue

        try:
            req = urllib.request.Request(img_url, headers={"User-Agent": "pptx-skill/1.0"})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                with open(filepath, "wb") as f:
                    f.write(resp.read())
            downloaded.append(filepath)
        except Exception as e:
            # 单次重试：网络波动常见，1秒后重试一次
            import time
            time.sleep(1)
            try:
                req = urllib.request.Request(img_url, headers={"User-Agent": "pptx-skill/1.0"})
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    with open(filepath, "wb") as f:
                        f.write(resp.read())
                downloaded.append(filepath)
            except Exception as e2:
                print(f"  [警告] 下载失败 id={img_id}: {e} (重试也失败: {e2})", file=sys.stderr)
                continue

    return downloaded


def search_and_download(
    query,
    count=3,
    output_dir="./images",
    size="webformat",
    **search_kwargs,
):
    """
    便捷函数：搜索 + 下载一步完成。

    返回:
        list[str]: 下载后的本地文件路径列表
    """
    hits = search_images(query, count=count, **search_kwargs)
    if not hits:
        print(f"  [提示] 未找到与 '{query}' 相关的图片", file=sys.stderr)
        return []
    print(f"  [信息] 找到 {len(hits)} 张图片，开始下载...")
    paths = download_images(hits, output_dir=output_dir, size=size)
    print(f"  [信息] 成功下载 {len(paths)} 张图片到 {os.path.abspath(output_dir)}")
    return paths


def _build_arg_parser():
    parser = argparse.ArgumentParser(
        description="从 Pixabay 搜索并下载免费可商用图片（供 PPT 制作使用）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python pixabay_search.py -q "artificial intelligence" -c 5 -o ./images
  python pixabay_search.py -q "nature" --category nature --orientation horizontal
  python pixabay_search.py -q "科技" --lang zh --image_type illustration
  python pixabay_search.py -q "business meeting" --min_width 1920 --min_height 1080
        """,
    )
    parser.add_argument("-q", "--query", required=True,
                        help="搜索关键词，多个词用空格或加号分隔")
    parser.add_argument("-c", "--count", type=int, default=3,
                        help="下载图片数量（默认 3，最多 20）")
    parser.add_argument("-o", "--output", default="./images",
                        help="图片保存目录（默认 ./images）")
    parser.add_argument("--api_key", default=DEFAULT_API_KEY,
                        help="Pixabay API key（默认使用内置 key）")
    parser.add_argument("--lang", default="zh",
                        choices=sorted(VALID_LANGS),
                        help="搜索语言（默认 zh）")
    parser.add_argument("--image_type", default="photo",
                        choices=sorted(VALID_IMAGE_TYPES),
                        help="图片类型（默认 photo）")
    parser.add_argument("--orientation", default="horizontal",
                        choices=sorted(VALID_ORIENTATIONS),
                        help="图片方向（默认 horizontal，适合 PPT）")
    parser.add_argument("--category", default=None,
                        choices=sorted(VALID_CATEGORIES),
                        help="分类过滤")
    parser.add_argument("--min_width", type=int, default=0,
                        help="最小宽度像素")
    parser.add_argument("--min_height", type=int, default=0,
                        help="最小高度像素")
    parser.add_argument("--colors", default=None,
                        help="颜色过滤，逗号分隔，如 red,blue")
    parser.add_argument("--editors_choice", default=None, type=str,
                        choices=["true", "false"],
                        help="仅获取编辑精选")
    parser.add_argument("--safesearch", default="true",
                        choices=["true", "false"],
                        help="安全搜索（默认 true）")
    parser.add_argument("--order", default="popular",
                        choices=sorted(VALID_ORDERS),
                        help="排序方式（默认 popular）")
    parser.add_argument("--page", type=int, default=1,
                        help="页码（默认 1）")
    parser.add_argument("--per_page", type=int, default=20,
                        help="每页结果数 3-200（默认 20）")
    parser.add_argument("--size", default="webformat",
                        choices=["preview", "webformat", "large", "fullHD"],
                        help="下载图片尺寸（默认 webformat，适合 PPT）")
    parser.add_argument("--json", action="store_true",
                        help="仅输出搜索结果 JSON，不下载图片")
    parser.add_argument("--timeout", type=int, default=15,
                        help="请求超时秒数（默认 15）")
    return parser


def main():
    parser = _build_arg_parser()
    args = parser.parse_args()

    # count 上限保护
    count = max(1, min(args.count, 20))

    editors_choice = None
    if args.editors_choice is not None:
        editors_choice = args.editors_choice == "true"

    try:
        hits = search_images(
            query=args.query,
            api_key=args.api_key,
            lang=args.lang,
            image_type=args.image_type,
            orientation=args.orientation,
            category=args.category,
            min_width=args.min_width,
            min_height=args.min_height,
            colors=args.colors,
            editors_choice=editors_choice,
            safesearch=args.safesearch == "true",
            order=args.order,
            page=args.page,
            per_page=args.per_page,
            count=count if not args.json else None,
            timeout=args.timeout,
        )
    except (ValueError, RuntimeError) as e:
        print(f"搜索失败: {e}", file=sys.stderr)
        sys.exit(1)

    if not hits:
        print(f"未找到与 '{args.query}' 相关的图片", file=sys.stderr)
        sys.exit(1)

    if args.json:
        print(json.dumps(hits, ensure_ascii=False, indent=2))
        return

    print(f"找到 {len(hits)} 张图片，开始下载到 {args.output} ...")
    paths = download_images(hits, output_dir=args.output, size=args.size, timeout=args.timeout)
    print(f"\n下载完成，成功 {len(paths)}/{len(hits)} 张：")
    for p in paths:
        print(f"  {os.path.abspath(p)}")


if __name__ == "__main__":
    main()
