#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Shapefile 属性表转 Excel（使用 pyogrio 直读，无需 geopandas/fiona）
用法：
  python SHP转EXCEL.py [shp路径] [-o 输出路径] [-e 编码] [-g]

依赖：pyogrio, pandas, openpyxl（均已安装）
"""

import pandas as pd
import pyogrio
import argparse
import sys
import os

# ================== 用户配置区 ==================
INPUT_SHP = r"E:\WMY2026\0701相城水稻\2026年7月17日15点36分\15点47分\相城区_水稻核查_2026_核查地块\320507000000_58_2026.shp"
OUTPUT_EXCEL = None          # None 自动生成，或指定完整 .xlsx 路径
ENCODING = "detect"          # "detect" 自动检测, "auto" 交互探测, 或手动指定如 "utf-8"
INCLUDE_GEOMETRY = False
# ================================================


def read_cpg_file(shp_path):
    """读取 .cpg 文件获取编码声明"""
    cpg_path = os.path.splitext(shp_path)[0] + '.cpg'
    if os.path.exists(cpg_path):
        try:
            with open(cpg_path, 'r', encoding='utf-8') as f:
                encoding = f.read().strip()
            if encoding:
                print(f"📋 .cpg 文件声明编码: {encoding}")
                return encoding
        except:
            pass
    return None


def detect_dbf_encoding(shp_path):
    """使用 chardet 检测 .dbf 文件编码"""
    dbf_path = os.path.splitext(shp_path)[0] + '.dbf'
    if not os.path.exists(dbf_path):
        return None

    try:
        import chardet
        with open(dbf_path, 'rb') as f:
            raw_data = f.read(min(100000, os.path.getsize(dbf_path)))
        result = chardet.detect(raw_data)
        encoding = result['encoding']
        confidence = result['confidence']
        print(f"🔍 chardet 检测结果: {encoding} (置信度: {confidence:.2%})")
        return encoding
    except ImportError:
        print("⚠️  chardet 未安装，跳过自动检测")
        print("   安装命令: pip install chardet")
        return None
    except Exception as e:
        print(f"⚠️  chardet 检测失败: {e}")
        return None


def probe_encodings(shp_path, encodings_list):
    """交互式探测编码：尝试多个编码，让用户肉眼选择"""
    print("\n🔍 开始交互式编码探测...")
    print("请根据显示的中文内容，判断哪个编码正确（中文应为可读汉字，非乱码）\n")

    results = []
    for enc in encodings_list:
        try:
            gdf = pyogrio.read_dataframe(shp_path, encoding=enc, read_geometry=False)
            # 找第一个文本字段作为样本
            text_cols = [c for c in gdf.columns if gdf[c].dtype.name == 'object']
            sample_col = text_cols[0] if text_cols else gdf.columns[0]
            sample_vals = [str(gdf[sample_col].iloc[i]).strip() for i in range(min(3, len(gdf)))]
            results.append((enc, True, sample_col, sample_vals))
            print(f"  ✅ {enc:15s} | 字段 '{sample_col}' 样本: {sample_vals}")
        except Exception as e:
            results.append((enc, False, None, None))
            print(f"  ❌ {enc:15s} | 读取失败: {e}")

    print("\n" + "="*60)
    print("请根据上面显示的中文内容，输入正确的编码编号：")
    print("="*60)

    valid_results = [(i, r) for i, r in enumerate(results) if r[1]]
    for idx, (_, (enc, _, _, vals)) in enumerate(valid_results, 1):
        print(f"  {idx}. {enc}")
        if vals:
            print(f"     样本: {vals}")

    while True:
        choice = input(f"\n请输入编号 (1-{len(valid_results)})，或直接输入编码名称: ").strip()
        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(valid_results):
                return valid_results[idx][1][0]
            print("❌ 无效编号，请重新选择。")
        else:
            try:
                pyogrio.read_dataframe(shp_path, encoding=choice, read_geometry=False)
                return choice
            except Exception as e:
                print(f"❌ 编码 '{choice}' 读取失败: {e}，请重新选择。")


def auto_detect_encoding(shp_path):
    """自动检测编码：优先 .cpg，再 chardet，最后尝试常见编码"""
    # 1. 读 .cpg 文件
    cpg_encoding = read_cpg_file(shp_path)
    if cpg_encoding:
        cpg_upper = cpg_encoding.upper().replace('-', '')
        if cpg_upper in ['UTF8', 'UTF_8']:
            cpg_encoding = 'utf-8'
        elif cpg_upper in ['GBK', 'CP936']:
            cpg_encoding = 'gbk'
        elif cpg_upper in ['GB18030']:
            cpg_encoding = 'gb18030'
        try:
            pyogrio.read_dataframe(shp_path, encoding=cpg_encoding, read_geometry=False)
            print(f"✅ 使用 .cpg 声明的编码: {cpg_encoding}")
            return cpg_encoding
        except:
            print(f"⚠️  .cpg 声明的编码 '{cpg_encoding}' 读取失败，继续检测...")

    # 2. chardet 检测
    detected = detect_dbf_encoding(shp_path)
    if detected:
        try:
            pyogrio.read_dataframe(shp_path, encoding=detected, read_geometry=False)
            print(f"✅ 使用 chardet 检测的编码: {detected}")
            return detected
        except:
            print(f"⚠️  chardet 检测的编码 '{detected}' 读取失败，继续尝试...")

    # 3. 逐个尝试常见编码 + 中文验证
    print("🔍 尝试常见编码...")
    common_encodings = ['utf-8', 'gb18030', 'gbk', 'gb2312', 'cp936', 'latin1']
    for enc in common_encodings:
        try:
            df = pyogrio.read_dataframe(shp_path, encoding=enc, read_geometry=False)
            for c in df.columns:
                if df[c].dtype.name == 'object':
                    sample = str(df[c].iloc[0]).strip()
                    if any('\u4e00' <= ch <= '\u9fff' for ch in sample):
                        print(f"✅ 自动选择编码: {enc} (检测到中文内容)")
                        return enc
                    break
            print(f"  ✓ {enc} 可读取，但未检测到中文，跳过")
        except:
            print(f"  ✗ {enc} 读取失败")

    print("⚠️  所有编码尝试失败，请使用 -e auto 交互探测，或手动指定编码")
    return None


def shp_to_excel(shp_path, excel_path=None, encoding='detect', include_geometry=False):
    """主转换函数"""

    # 确定编码
    used_enc = None
    if encoding.lower() == 'auto':
        encodings_to_try = ['utf-8', 'gb18030', 'gbk', 'gb2312', 'cp936', 'latin1', 'iso-8859-1']
        used_enc = probe_encodings(shp_path, encodings_to_try)
    elif encoding.lower() == 'detect':
        used_enc = auto_detect_encoding(shp_path)
        if used_enc is None:
            print("\n❌ 自动检测失败，请使用 -e auto 进行交互探测，或手动指定编码")
            sys.exit(1)
    else:
        used_enc = encoding

    print(f"\n📖 使用编码 '{used_enc}' 读取 Shapefile...")

    try:
        gdf = pyogrio.read_dataframe(shp_path, encoding=used_enc, read_geometry=not include_geometry)
        print(f"✅ 成功读取 {len(gdf)} 条记录")
    except Exception as e:
        print(f"❌ 读取失败: {e}")
        sys.exit(1)

    # 显示字段信息
    print(f"\n📊 字段列表 ({len(gdf.columns)} 个):")
    for c in gdf.columns:
        sample = str(gdf[c].iloc[0]).strip()[:30]
        print(f"  - {c:20s} 样本: {sample}")

    # 输出路径
    if excel_path is None:
        excel_path = os.path.splitext(shp_path)[0] + '.xlsx'

    out_dir = os.path.dirname(excel_path)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir)
        print(f"📁 已创建输出目录: {out_dir}")

    print(f"\n💾 正在导出 Excel...")
    try:
        gdf.to_excel(excel_path, index=False, engine='openpyxl')
        print(f"✅ 成功导出到: {excel_path}")
        print(f"📌 使用编码: {used_enc}")
        print(f"📊 共 {len(gdf)} 行 × {len(gdf.columns)} 列")
    except Exception as e:
        print(f"❌ 导出 Excel 失败: {e}")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description='Shapefile 属性表转 Excel（自动编码检测）',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
编码说明：
  -e detect   自动检测（默认，优先读 .cpg 文件，再用 chardet）
  -e auto     交互式探测，让用户肉眼选择
  -e utf-8    手动指定编码

示例：
  python SHP转EXCEL.py                         # 自动检测
  python SHP转EXCEL.py -e auto                  # 交互探测
  python SHP转EXCEL.py -o output.xlsx -g        # 含几何列
        """)
    parser.add_argument('shp', nargs='?', default=INPUT_SHP, help='输入的 Shapefile 路径')
    parser.add_argument('-o', '--output', default=OUTPUT_EXCEL, help='输出的 Excel 文件路径')
    parser.add_argument('-e', '--encoding', default=ENCODING,
                        help='编码: detect(自动), auto(交互), 或具体编码名')
    parser.add_argument('-g', '--include-geometry', action='store_true', default=INCLUDE_GEOMETRY,
                        help='包含几何列（WKT格式）')
    args = parser.parse_args()

    if not os.path.exists(args.shp):
        print(f"❌ 文件不存在: {args.shp}")
        sys.exit(1)

    print("="*60)
    print("  Shapefile 属性表 → Excel 转换工具")
    print("="*60)
    print(f"输入文件: {args.shp}")

    shp_to_excel(args.shp, args.output, args.encoding, args.include_geometry)

    print("\n" + "="*60)
    print("  转换完成！")
    print("="*60)


if __name__ == '__main__':
    main()
