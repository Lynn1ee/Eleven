#!/usr/bin/env python
r"""
客服数据周报 HTML -> PNG 截图工具

用法:
    python screenshot_report.py <HTML路径> [PNG输出路径]

示例:
    python screenshot_report.py report.html
    python screenshot_report.py report.html report.png
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright


def screenshot(html_path, png_path, width=1200):
    html_path = Path(html_path).resolve()
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={'width': width, 'height': 800})
        page.goto(html_path.as_uri())
        page.wait_for_timeout(3000)  # 等待 Chart.js 渲染
        height = page.evaluate('() => document.body.scrollHeight')
        page.set_viewport_size({'width': width, 'height': height + 20})
        page.wait_for_timeout(500)
        page.screenshot(path=str(png_path), full_page=True)
        browser.close()
    return png_path


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    html_path = sys.argv[1]
    src = Path(html_path)

    if len(sys.argv) >= 3:
        png_path = Path(sys.argv[2])
    else:
        png_path = src.with_suffix('.png')

    out = screenshot(html_path, png_path)

    size_kb = png_path.stat().st_size / 1024
    print(f'[OK] PNG 截图已保存: {out}')
    print(f'     文件大小: {size_kb:.0f} KB')


if __name__ == '__main__':
    main()
