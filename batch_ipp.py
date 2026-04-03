"""
WOW English 知识产权投诉 - 批量处理工具 v1.0
输入：Excel 链接列表 → 自动抓取 → 生成 IPP 投诉文档

使用方法：
  1. 打开「投诉链接模板.xlsx」，在"商品链接"列填入要投诉的链接（一行一个）
  2. 双击运行「批量处理.bat」，或命令行运行：python batch_ipp.py
  3. 处理完成后，打开 output/IPP投诉文档_日期.xlsx，复制粘贴到 IPP 平台
"""

import os
import sys
import io
import re
import time
import random
import socket
from datetime import datetime
from pathlib import Path

# Windows 编码修复
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    try:
        os.system('chcp 65001 >nul 2>&1')
    except:
        pass

try:
    from playwright.sync_api import sync_playwright, Page
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    import easyocr
except ImportError as e:
    print(f"[ERROR] 缺少依赖: {e}")
    print("运行: pip install playwright openpyxl easyocr")
    sys.exit(1)

# ==================== 配置 ====================
BRAND_KEYWORDS = [
    "wow english", "wowenglish", "wow english",
    "史蒂夫英语", "steve english",
    "english singsing", "englishsingsing",
    "英语启蒙动画", "史蒂夫", "steve",
    "wowyoung", "wow young",
]
COPYRIGHT_HOLDER = "WOW English 版权方"
RIGHTS_TYPE = "商标权/著作权"
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")

# ==================== OCR 预加载 ====================
print("[OCR] 初始化文字识别引擎（首次约30秒）...")
_OCR_READER = None
def get_ocr_reader():
    global _OCR_READER
    if _OCR_READER is None:
        _OCR_READER = easyocr.Reader(['ch_sim', 'en'], gpu=False, verbose=False)
    return _OCR_READER


# ==================== 平台检测 ====================
def detect_platform(url: str) -> str:
    if 'jd.com' in url:
        return 'jd'
    elif 'tmall.com' in url or 'taobao.com' in url:
        return 'tmall'
    elif 'pinduoduo.com' in url or 'pdd.com' in url:
        return 'pdd'
    return 'unknown'


# ==================== 京东提取器 ====================
class JDExtractor:
    TITLE_SELECTORS = [
        '.product-title__main', '[class*="product-title"]',
        '[class*="sku-name"]', '[class*="item-name"]', '.p-title', '#itemName',
        'div[class*="product-name"] h1', '#name h1', 'h1[class*="name"]',
        'h1[class*="product"]', 'h1[class*="item"]', 'h1', '[data-sku]',
    ]
    PRICE_SELECTORS = [
        '[class*="price"] [class*="integer"]', '[class*="price"] strong',
        '.p-price .price', '#jdprice', '[data-price]', '[class*="sale-price"]',
        '[class*="price-wrap"] span', '.J-p-', '#spec-qty-price',
    ]
    SHOP_SELECTORS = [
        '[class*="shop-name"] a[href*="shop"]', '.j-shop-content .shop-name',
        '#shopInfo .shop-name', '#shop-name',
        '[class*="seller-info"] [class*="name"]',
        '#pop_shop .name', '.popshop .shop-name',
        '[class*="shop-name"] a', '.shop-name a', '[class*="seller"] a',
        '[class*="store"] a', '[class*="shop"]',
    ]

    @classmethod
    def extract(cls, page: Page, url: str) -> dict:
        data = {
            "platform": "京东",
            "title": "", "price": "", "shop_name": "", "seller_id": "",
            "sales": "", "main_images": [], "url": url,
            "infringement_check": "否"
        }
        js_data = cls._extract_via_js(page)
        if js_data:
            data.update(js_data)
        else:
            data["title"] = cls._try_selectors(page, cls.TITLE_SELECTORS)
            data["price"] = cls._try_selectors(page, cls.PRICE_SELECTORS)
            data["shop_name"] = cls._try_selectors(page, cls.SHOP_SELECTORS)
        data["infringement_check"] = cls._check_infringement(data)
        return data

    @staticmethod
    def _extract_via_js(page: Page) -> dict:
        try:
            result = page.evaluate("""
                () => {
                    const data = {};
                    const noiseList = ['最小单价','搭配购买','似乎出了点问题','先看看','极速审核',
                                       '立即购买','加入购物车','查看全部','售后服务','退换无忧'];
                    function isNoise(t) { return noiseList.some(n => t.includes(n)); }

                    const h1 = document.querySelector('h1') || document.querySelector('[class*="product-title"]');
                    if (h1) { const t = h1.innerText.trim().substring(0,200); if(t && t.length>5 && !isNoise(t)) data.title = t; }

                    const priceEls = document.querySelectorAll('[class*="price"]');
                    for(const el of priceEls){ const t=el.innerText.trim(); if(/^￥|^¥|^\\d/.test(t)&&!isNoise(t)&&t.length<30){data.price=t;break;} }

                    const shopSelectors = [
                        '[class*="shop-info"] [class*="name"]','[class*="shop-header"] [class*="name"]',
                        '[class*="seller-info"] [class*="name"]','[class*="score"] [class*="shop"]',
                        '[class*="shop-name"]','a[href*="shop.jd.com"]','[class*="shop-name"] a','#shop-name a'
                    ];
                    for(const sel of shopSelectors){
                        const el=document.querySelector(sel);
                        if(el){const t=el.innerText.trim();if(t&&!isNoise(t)&&t.length>1&&t.length<50){data.shop_name=t;break;}}
                    }
                    const shopLink = document.querySelector('a[href*="shop.jd.com"]')||document.querySelector('[class*="shop-name"] a');
                    if(shopLink&&shopLink.href){
                        const m=shopLink.href.match(/shopId[=:]"?(\\d+)/)||shopLink.href.match(/\\/shop\\/(\\d+)/);
                        if(m) data.seller_id=m[1];
                    }
                    const salesEl = document.querySelector('.p-commit a')||document.querySelector('#commit_count')||document.querySelector('.J-comm-count');
                    if(salesEl){const t=salesEl.innerText.trim().split('\\n')[0];if(t&&!isNoise(t)&&t.length<30)data.sales=t;}
                    return data;
                }
            """)
            return {k: v for k, v in result.items() if v}
        except:
            return {}

    @staticmethod
    def _try_selectors(page, selectors):
        for sel in selectors:
            try:
                el = page.locator(sel).first
                if el.count() > 0 and el.is_visible():
                    text = el.inner_text().strip()
                    if text and len(text) > 1:
                        return text[:200]
            except:
                continue
        return ""

    @staticmethod
    def _check_infringement(data):
        text = f"{data.get('title','')} {data.get('shop_name','')}".lower()
        matched = [kw for kw in BRAND_KEYWORDS if kw.lower() in text]
        return f"是（匹配：{', '.join(matched)}）" if matched else "否"

    @staticmethod
    def ocr_shop_name(image_path):
        try:
            from PIL import Image
            reader = get_ocr_reader()
            img = Image.open(image_path)
            w, h = img.size
            top = img.crop((0, 0, w, int(h * 0.25)))
            tmp = image_path.replace('.png','_shop.png')
            top.save(tmp)
            results = reader.readtext(tmp)
            all_text = ' '.join([r[1] for r in results])
            shop_match = re.search(
                r'([\u4e00-\u9fa5]{2,30}(?:旗舰店|专营店|专卖店|直营店|官方旗舰店|海外旗舰店|'
                r'官方店|商城|百货|书屋|书店))', all_text)
            if shop_match:
                return shop_match.group(1).strip()
            chinese = re.findall(r'[\u4e00-\u9fa5]{3,}', all_text)
            if chinese:
                return max(chinese, key=len)
        except:
            pass
        return ""


# ==================== 拼多多提取器 ====================
class PDDExtractor:
    @classmethod
    def extract(cls, page: Page, url: str) -> dict:
        data = {
            "platform": "拼多多",
            "title": "", "price": "", "shop_name": "", "seller_id": "",
            "sales": "", "main_images": [], "url": url,
            "infringement_check": "否"
        }
        js_data = cls._extract_via_js(page)
        if js_data:
            data.update(js_data)
        data["infringement_check"] = JDExtractor._check_infringement(data)
        return data

    @staticmethod
    def _extract_via_js(page: Page) -> dict:
        try:
            result = page.evaluate("""
                () => {
                    const data = {};
                    const noiseList = ['最小单价','搭配购买','极速审核','立即购买','加入购物车','查看全部'];
                    function isNoise(t) { return noiseList.some(n => t.includes(n)); }

                    const titleEl = document.querySelector('[class*="product-title"]') ||
                                    document.querySelector('[class*="goods-title"]') ||
                                    document.querySelector('h1') || document.querySelector('[class*="item-title"]');
                    if(titleEl){const t=titleEl.innerText.trim().substring(0,200);if(t&&t.length>3&&!isNoise(t))data.title=t;}

                    const priceEl = document.querySelector('[class*="price"] [class*="value"]') ||
                                   document.querySelector('[class*="sale-price"]') ||
                                   document.querySelector('.product-price') || document.querySelector('[class*="price"]');
                    if(priceEl){const t=priceEl.innerText.trim().substring(0,20);if(t&&!isNoise(t))data.price=t;}

                    const shopEl = document.querySelector('[class*="shop-name"]') ||
                                  document.querySelector('[class*="store-name"]') ||
                                  document.querySelector('[class*="malls-name"]') ||
                                  document.querySelector('[class*="goods-single"]');
                    if(shopEl){const t=shopEl.innerText.trim().substring(0,50);if(t&&!isNoise(t))data.shop_name=t;}

                    const salesEl = document.querySelector('[class*="sales"]') || document.querySelector('[class*="sales-count"]');
                    if(salesEl){const t=salesEl.innerText.trim().substring(0,20);if(t&&!isNoise(t))data.sales=t;}

                    return data;
                }
            """)
            return {k: v for k, v in result.items() if v}
        except:
            return {}


# ==================== 天猫/淘宝提取器 ====================
class TMExtractor:
    @classmethod
    def extract(cls, page: Page, url: str) -> dict:
        data = {
            "platform": "天猫/淘宝",
            "title": "", "price": "", "shop_name": "", "seller_id": "",
            "sales": "", "main_images": [], "url": url,
            "infringement_check": "否"
        }
        js_data = cls._extract_via_js(page)
        if js_data:
            data.update(js_data)
        data["infringement_check"] = JDExtractor._check_infringement(data)
        return data

    @staticmethod
    def _extract_via_js(page: Page) -> dict:
        try:
            result = page.evaluate("""
                () => {
                    const data = {};
                    const noiseList = ['最小单价','搭配购买','极速审核','立即购买','加入购物车'];
                    function isNoise(t) { return noiseList.some(n => t.includes(n)); }

                    const titleEl = document.querySelector('[class*="product-title"]') ||
                                    document.querySelector('.product-title') ||
                                    document.querySelector('h1[class*="title"]') || document.querySelector('h1');
                    if(titleEl){const t=titleEl.innerText.trim().substring(0,200);if(t&&t.length>3&&!isNoise(t))data.title=t;}

                    const priceEl = document.querySelector('#J_StrPrice') || document.querySelector('[class*="price"]');
                    if(priceEl){const t=priceEl.innerText.trim().substring(0,20);if(t)data.price=t;}

                    const shopEl = document.querySelector('[class*="shop-name"] a') ||
                                  document.querySelector('#shopExtra a') ||
                                  document.querySelector('[class*="shopInfo"]');
                    if(shopEl){const t=shopEl.innerText.trim().substring(0,50);if(t&&!isNoise(t))data.shop_name=t;}

                    return data;
                }
            """)
            return {k: v for k, v in result.items() if v}
        except:
            return {}


# ==================== 截图管理 ====================
class ScreenshotManager:
    def __init__(self, page: Page, save_dir: str, prefix: str = ""):
        self.page = page
        self.save_dir = save_dir
        self.prefix = prefix
        os.makedirs(save_dir, exist_ok=True)
        os.makedirs(os.path.join(save_dir, "main_images"), exist_ok=True)

    def screenshot_page(self) -> str:
        filename = os.path.join(self.save_dir, f"{self.prefix}screenshot.png")
        h = self.page.evaluate("document.body.scrollHeight")
        self.page.set_viewport_size({"width": 1366, "height": min(h, 6000)})
        self.page.screenshot(path=filename, full_page=True)
        return filename

    def download_main_images(self, urls: list) -> list:
        saved = []
        for i, url in enumerate(urls[:5]):
            try:
                url_clean = re.sub(r'/\d+x\d+/', '/800x800/', url)
                filename = os.path.join(self.save_dir, "main_images", f"{self.prefix}img_{i+1}.jpg")
                resp = self.page.context.request.get(url_clean)
                if resp.ok:
                    with open(filename, 'wb') as f:
                        f.write(resp.body())
                    saved.append(filename)
            except:
                continue
        return saved


# ==================== IPP 投诉文档生成器 ====================
class IPPComplaintGenerator:
    """生成 IPP 平台可直接粘贴的投诉文档"""

    def __init__(self, results: list, save_dir: str):
        self.results = results
        self.save_dir = save_dir
        os.makedirs(save_dir, exist_ok=True)

    def generate_excel(self) -> str:
        """生成 IPP 投诉 Excel（结构化格式，可复制粘贴）"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "IPP投诉文档"

        # 标题行
        title_fill = PatternFill(start_color="C41E3A", end_color="C41E3A", fill_type="solid")
        title_font = Font(bold=True, color="FFFFFF", size=12)
        header_fill = PatternFill(start_color="FFD0D0", end_color="FFD0D0", fill_type="solid")
        header_font = Font(bold=True, size=10)

        # 表头
        headers = [
            "序号", "平台", "商品链接", "商品标题", "店铺名称",
            "卖家ID", "价格", "销量/评价", "疑似侵权", "截图路径", "证据目录"
        ]
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col, value=h)
            c.fill = title_fill
            c.font = title_font
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = self._thin_border()

        # 维权方信息（固定内容，第2行）
        rights_info = {
            "序号": "", "平台": "维权方",
            "商品链接": "权利人/公司", "商品标题": COPYRIGHT_HOLDER,
            "店铺名称": "维权类型", "卖家ID": RIGHTS_TYPE,
            "价格": "处理平台", "销量/评价": "京东/淘宝/拼多多",
            "疑似侵权": "投诉原因", "截图路径": "请参考附件截图",
            "证据目录": "见对应子文件夹"
        }
        row2_fill = PatternFill(start_color="FFF0F0", end_color="FFF0F0", fill_type="solid")
        for col, h in enumerate(headers, 1):
            val = rights_info.get(h, "")
            c = ws.cell(row=2, column=col, value=val)
            c.fill = row2_fill
            c.font = Font(size=9, italic=True)
            c.border = self._thin_border()

        # 数据行
        for i, r in enumerate(self.results, 3):
            is_infringe = "是" in r.get("infringement_check", "否")
            row_fill = PatternFill(
                start_color="FFCCCC" if is_infringe else "FFFFFF",
                end_color="FFCCCC" if is_infringe else "FFFFFF",
                fill_type="solid"
            )
            row_data = [
                str(i - 2),
                r.get("platform", ""),
                r.get("url", ""),
                r.get("title", ""),
                r.get("shop_name", ""),
                r.get("seller_id", ""),
                r.get("price", ""),
                r.get("sales", ""),
                r.get("infringement_check", "否"),
                r.get("screenshot_path", ""),
                r.get("save_dir", ""),
            ]
            for col, val in enumerate(row_data, 1):
                c = ws.cell(row=i, column=col, value=val)
                c.fill = row_fill
                c.alignment = Alignment(wrap_text=True, vertical="top")
                c.border = self._thin_border()
                if col == 9 and is_infringe:
                    c.font = Font(bold=True, color="CC0000")
                else:
                    c.font = Font(size=9)

        # 列宽
        widths = [6, 10, 55, 50, 25, 18, 12, 15, 35, 40, 50]
        for ci, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w

        ws.row_dimensions[1].height = 30

        date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(self.save_dir, f"IPP投诉文档_{date_str}.xlsx")
        wb.save(filename)
        return filename

    def generate_txt(self) -> str:
        """生成纯文本版 IPP 投诉文档（直接复制粘贴用）"""
        date_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        lines = []
        lines.append("=" * 70)
        lines.append("           WOW English 知识产权投诉文档")
        lines.append("           生成时间：" + date_str)
        lines.append("=" * 70)
        lines.append("")
        lines.append("【权利人信息】")
        lines.append(f"  权利人：{COPYRIGHT_HOLDER}")
        lines.append(f"  维权类型：{RIGHTS_TYPE}")
        lines.append(f"  处理平台：京东 / 天猫 / 拼多多")
        lines.append("")
        lines.append("=" * 70)
        lines.append("【投诉商品列表】")
        lines.append("=" * 70)

        for i, r in enumerate(self.results, 1):
            lines.append("")
            lines.append(f"─── 商品 {i} ───")
            lines.append(f"  平台：{r.get('platform','')}")
            lines.append(f"  商品链接：{r.get('url','')}")
            lines.append(f"  商品标题：{r.get('title','')}")
            lines.append(f"  店铺名称：{r.get('shop_name','')}")
            lines.append(f"  卖家ID：{r.get('seller_id','')}")
            lines.append(f"  价格：{r.get('price','')}")
            lines.append(f"  销量：{r.get('sales','')}")
            lines.append(f"  疑似侵权：{r.get('infringement_check','')}")
            lines.append(f"  截图：{r.get('screenshot_path','')}")
            lines.append(f"  证据目录：{r.get('save_dir','')}")

        lines.append("")
        lines.append("=" * 70)
        lines.append("【投诉理由模板（可直接复制到 IPP 平台）】")
        lines.append("=" * 70)

        for i, r in enumerate(self.results, 1):
            shop = r.get('shop_name', '未知店铺')
            title = r.get('title', '未知商品')
            brand = COPYRIGHT_HOLDER
            lines.append(f"""
商品{i}投诉理由：
您好，我方是「{brand}」的版权方。
经调查发现，商家「{shop}」销售的商品标题/描述中使用了 WOW English 相关品牌关键词，
涉嫌构成商标侵权及著作权侵权。
商品链接：{r.get('url','')}
商品标题：{title}
请贵平台依据《电子商务法》第42-45条及《商标法》相关规定，
对上述侵权商品采取删除、屏蔽、断开链接等必要措施。
如有需要，我方可提供完整著作权登记证书及商标授权文件。
""")

        lines.append("=" * 70)
        lines.append(f"共 {len(self.results)} 件商品 | 生成时间：{date_str}")
        lines.append("=" * 70)

        content = '\n'.join(lines)
        txt_path = os.path.join(self.save_dir, f"IPP投诉文档_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return txt_path

    def generate_batch_summary(self) -> str:
        """生成批量处理汇总 Excel"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "汇总"

        # 样式
        hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        hdr_font = Font(bold=True, color="FFFFFF", size=11)
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")

        headers = ["序号", "平台", "商品链接", "商品标题（已脱敏）",
                   "店铺名称", "价格", "侵权判定", "截图文件", "证据目录"]
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col, value=h)
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal="center")

        for i, r in enumerate(self.results, 2):
            is_infringe = "是" in r.get("infringement_check", "否")
            # 标题脱敏：只显示前20字+***
            title_raw = r.get('title', '')[:20]
            title_display = title_raw + ('...' if len(r.get('title','')) > 20 else '')
            # 链接脱敏
            url_raw = r.get('url', '')
            platform = r.get('platform', '')
            row_data = [
                str(i - 1),
                platform,
                url_raw,
                title_display,
                r.get('shop_name', ''),
                r.get('price', ''),
                r.get('infringement_check', '否'),
                os.path.basename(r.get('screenshot_path', '')),
                os.path.basename(r.get('save_dir', '')),
            ]
            for col, val in enumerate(row_data, 1):
                c = ws.cell(row=i, column=col, value=val)
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if col == 7:
                    c.fill = red_fill if is_infringe else green_fill
                    c.font = Font(bold=True, color="FFFFFF" if is_infringe else "000000")

        widths = [6, 10, 50, 28, 25, 12, 35, 25, 25]
        for ci, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w

        # 统计行
        last_row = len(self.results) + 2
        ws.cell(row=last_row, column=1, value="统计").font = Font(bold=True)
        ws.cell(row=last_row, column=2, value=f"共 {len(self.results)} 件")
        ws.cell(row=last_row, column=3, value=f"疑似侵权 {sum(1 for r in self.results if '是' in r.get('infringement_check',''))} 件")

        date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(self.save_dir, f"批量处理汇总_{date_str}.xlsx")
        wb.save(filename)
        return filename

    @staticmethod
    def _thin_border():
        thin = Side(style='thin', color="AAAAAA")
        return Border(left=thin, right=thin, top=thin, bottom=thin)


# ==================== Excel 链接读取器 ====================
def read_links_from_excel(excel_path: str) -> list:
    """从 Excel 读取链接列表"""
    wb = openpyxl.load_workbook(excel_path)
    links = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell and isinstance(cell, str):
                    url = cell.strip()
                    if url and ('jd.com' in url or 'tmall.com' in url or
                                'taobao.com' in url or 'pinduoduo.com' in url or 'pdd.com' in url):
                        links.append(url)
    return links


# ==================== CDP 端口检测 ====================
def find_cdp_port():
    for port in [9222, 9223, 9224]:
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            r = sock.connect_ex(('127.0.0.1', port))
            sock.close()
            if r == 0:
                return port
        except:
            continue
    return None


# ==================== 拟人化延迟 ====================
def human_delay(min_sec: float = 0.5, max_sec: float = 2.0):
    """模拟人类操作的随机延迟"""
    t = random.uniform(min_sec, max_sec)
    time.sleep(t)

def human_scroll(page, steps: int = None):
    """拟人化滚动页面"""
    if steps is None:
        steps = random.randint(3, 6)
    for i in range(steps):
        y = i * random.randint(500, 900)
        page.evaluate(f"window.scrollTo(0, {y})")
        human_delay(0.4, 1.2)


# ==================== 单个链接抓取 ====================
def process_single_url(browser, url: str, batch_dir: str, idx: int, total: int) -> dict:
    """处理单个链接，返回提取结果"""
    cdp_port = find_cdp_port()
    platform = detect_platform(url)
    safe_name = re.sub(r'[^\w\u4e00-\u9fa5]', '_', url.split('/')[-1][:30])
    save_dir = os.path.join(batch_dir, f"{idx:03d}_{platform}_{safe_name}")
    os.makedirs(save_dir, exist_ok=True)

    print(f"\n[{idx}/{total}] 处理中... [{platform}] {url[:60]}...")

    # 随机延迟：进入下一个链接前等一等
    if idx > 1:
        wait = random.uniform(4, 12)
        print(f"  [💤] 随机等待 {wait:.1f} 秒...")
        time.sleep(wait)

    with sync_playwright() as p:
        page = None
        ctx = None

        if cdp_port:
            try:
                browser_ref = p.chromium.connect_over_cdp(f"http://127.0.0.1:{cdp_port}")
                ctx = browser_ref.contexts[0]
                page = ctx.new_page()
            except:
                pass

        if page is None:
            br = p.chromium.launch(
                headless=False,
                args=['--disable-blink-features=AutomationControlled',
                      '--disable-dev-shm-usage', '--no-sandbox']
            )
            ctx = br.new_context(
                viewport={"width": 1366, "height": 768},
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
                locale='zh-CN'
            )
            page = ctx.new_page()

        # 打开新标签页前随机等一下
        human_delay(0.3, 0.8)

        # 导航
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=30000)
            human_delay(1.5, 4.0)  # 页面 DOM 加载后等一下
            page.wait_for_load_state("networkidle", timeout=15000)
        except:
            try:
                page.goto(url, wait_until="load", timeout=30000)
                human_delay(2.0, 5.0)
            except Exception as e:
                print(f"  [WARN] 页面加载失败: {e}")

        # 模拟人类阅读页面的滚动
        human_delay(1.0, 3.0)
        human_scroll(page)

        # 提取
        if platform == 'jd':
            extractor = JDExtractor
        elif platform == 'pdd':
            extractor = PDDExtractor
        else:
            extractor = TMExtractor

        data = extractor.extract(page, url)

        # 截图
        shot_mgr = ScreenshotManager(page, save_dir, prefix=f"{idx:03d}_")
        screenshot_path = shot_mgr.screenshot_page()

        # OCR 兜底
        if platform == 'jd' and (not data.get('shop_name') or data['shop_name'] == '未提取'):
            ocr_shop = extractor.ocr_shop_name(screenshot_path) if hasattr(extractor, 'ocr_shop_name') else ''
            if ocr_shop:
                data['shop_name'] = ocr_shop
                print(f"  [OCR] 识别店铺名: {ocr_shop}")

        # 截图文件（给 IPP 用）
        data['screenshot_path'] = screenshot_path
        data['save_dir'] = save_dir

        # 关闭
        try:
            ctx.browser.close()
        except:
            pass

    return data


# ==================== 主程序 ====================
def main():
    print("""
╔══════════════════════════════════════════════════════════╗
║      WOW English 知识产权投诉 - 批量处理工具 v1.0         ║
║      📋 Excel 填链接 → 自动抓取 → 生成 IPP 投诉文档       ║
╚══════════════════════════════════════════════════════════╝
    """)

    # 读取 Excel 中的链接
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, "投诉链接模板.xlsx")

    excel_path = None
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    elif os.path.exists(template_path):
        excel_path = template_path
        print(f"[INPUT] 自动找到模板文件：{template_path}")
    else:
        print("[INPUT] 请提供 Excel 文件路径，或将文件命名为「投诉链接模板.xlsx」放在同目录")
        excel_path = input(">>> 拖入或输入路径：").strip().strip('"')

    if not os.path.exists(excel_path):
        print(f"[ERROR] 文件不存在：{excel_path}")
        print("提示：将您的链接 Excel 命名为「投诉链接模板.xlsx」放在脚本同目录即可")
        return

    links = read_links_from_excel(excel_path)
    if not links:
        print("[ERROR] 未从 Excel 中找到任何商品链接！")
        print("请确保 Excel 中有 jd.com / tmall.com / taobao.com / pinduoduo.com 链接")
        return

    print(f"\n[INFO] 共找到 {len(links)} 个链接：")
    for i, lnk in enumerate(links, 1):
        print(f"  {i}. {lnk[:70]}")
    print()

    # 创建批次输出目录
    batch_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    batch_dir = os.path.join(OUTPUT_DIR, f"batch_{batch_ts}")
    os.makedirs(batch_dir, exist_ok=True)

    # Chrome 调试端口提示
    cdp_port = find_cdp_port()
    if not cdp_port:
        print("""
[⚠️  提示] 未检测到 Chrome 调试模式！
建议先双击「启动已登录Chrome.bat」，再运行本脚本。
或直接在浏览器登录电商平台，脚本会处理验证码。
按回车继续（需手动处理验证码）...
        """)
        input()

    # 批量处理
    results = []
    failed = []

    for i, url in enumerate(links, 1):
        try:
            result = process_single_url(None, url, batch_dir, i, len(links))
            results.append(result)

            status = "✅" if "是" in result.get("infringement_check", "否") else "⚪"
            print(f"  {status} 标题：{result.get('title','')[:40]}")
            print(f"  {status} 店铺：{result.get('shop_name','')}")
            print(f"  {status} 侵权：{result.get('infringement_check','')}")

            # 每处理完一个，随机等待（模拟人类操作节奏）
            gap = random.uniform(3, 10)
            print(f"  [💤] 间隔 {gap:.1f} 秒...")
            time.sleep(gap)
        except Exception as e:
            print(f"  ❌ 处理失败: {e}")
            failed.append({"url": url, "error": str(e)})
            results.append({
                "platform": detect_platform(url), "title": "提取失败",
                "url": url, "shop_name": "", "seller_id": "",
                "price": "", "sales": "", "infringement_check": f"失败：{e}",
                "screenshot_path": "", "save_dir": batch_dir
            })

    # 生成 IPP 文档
    print(f"\n\n{'='*60}")
    print("[📄] 生成 IPP 投诉文档...")

    generator = IPPComplaintGenerator(results, batch_dir)

    excel_file = generator.generate_excel()
    print(f"  ✅ {excel_file}")

    txt_file = generator.generate_txt()
    print(f"  ✅ {txt_file}")

    summary_file = generator.generate_batch_summary()
    print(f"  ✅ {summary_file}")

    # 统计
    total = len(results)
    infr = sum(1 for r in results if '是' in r.get('infringement_check', ''))
    fail = len(failed)

    print(f"""
╔══════════════════════════════════════════════════════════╗
║                   ✅ 批量处理完成！                        ║
╚══════════════════════════════════════════════════════════╝

【统计】
  📦 总计处理：{total} 件
  🔴 疑似侵权：{infr} 件
  ⚠️  处理失败：{fail} 件

【输出目录】
  {batch_dir}

【IPP 投诉文档使用方式】
  1. 打开「IPP投诉文档_*.xlsx」
  2. 复制对应单元格内容到 IPP 平台表单
  3. 截图证据在对应编号的子文件夹中
  4. TXT 文件可直接全选复制到投诉理由文本框

【文件说明】
  📊 IPP投诉文档_*.xlsx  → IPP 平台表单（主要使用这个）
  📄 IPP投诉文档_*.txt   → 投诉理由模板（可直接粘贴）
  📈 批量处理汇总_*.xlsx → 处理结果总览表
""")


if __name__ == "__main__":
    main()
