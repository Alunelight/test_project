#!/usr/bin/env python3
"""
PDF文件匹配移动脚本（按姓名匹配）
根据Excel文件中的姓名匹配PDF文件，并将匹配的文件移动到"匹配结果"文件夹
"""

import re
import argparse
import shutil
from pathlib import Path
from typing import Set, Optional
import pandas as pd
from typing import Literal


def read_excel_dataframe(excel_path: Path) -> tuple[pd.DataFrame, str]:
    """
    读取Excel文件，返回DataFrame和姓名列名

    Args:
        excel_path: Excel文件路径

    Returns:
        (DataFrame, 姓名列名)元组

    Raises:
        FileNotFoundError: Excel文件不存在
        ValueError: Excel文件格式错误或找不到姓名列
    """
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel文件不存在: {excel_path}")

    try:
        # 根据文件扩展名选择首选引擎
        file_ext = excel_path.suffix.lower()
        if file_ext == ".xls":
            preferred_engines: list[Literal["xlrd", "openpyxl"]] = ["xlrd", "openpyxl"]
        elif file_ext in [".xlsx", ".xlsm"]:
            preferred_engines = ["openpyxl", "xlrd"]
        else:
            preferred_engines = ["xlrd", "openpyxl"]

        # 尝试使用引擎读取文件
        df = None
        last_error = None
        for engine in preferred_engines:
            try:
                df = pd.read_excel(excel_path, engine=engine)
                break
            except Exception as e:
                last_error = e
                continue

        if df is None:
            raise ValueError(f"无法读取Excel文件: {last_error}")

        # 查找"姓名"列
        name_col = None
        for col in df.columns:
            col_str = str(col).strip()
            if "姓名" in col_str:
                name_col = col
                break

        if name_col is None:
            raise ValueError(f"在Excel文件中找不到'姓名'列。可用列: {list(df.columns)}")

        return (df, name_col)

    except Exception as e:
        if isinstance(e, (FileNotFoundError, ValueError)):
            raise
        raise ValueError(f"读取Excel文件时出错: {str(e)}")


def read_excel_names(excel_path: Path) -> Set[str]:
    """
    读取Excel文件，提取姓名列的所有值

    Args:
        excel_path: Excel文件路径

    Returns:
        姓名集合

    Raises:
        FileNotFoundError: Excel文件不存在
        ValueError: Excel文件格式错误或找不到姓名列
    """
    df, name_col = read_excel_dataframe(excel_path)

    # 提取所有姓名，去除空值和NaN
    names = set()
    for idx, row in df.iterrows():
        name_value = str(row[name_col]).strip()
        # 跳过空值和NaN
        if pd.isna(row[name_col]) or name_value == "nan" or name_value == "":
            continue
        names.add(name_value)

    return names


def extract_name_from_filename(pdf_filename: str) -> Optional[str]:
    """
    从PDF文件名中提取员工姓名

    支持的文件名格式：
    - 陈玲-承诺书.pdf
    - 承诺书-陈冬如.pdf
    - 吴慧贤-承诺书(2).pdf
    - 承诺书-姓名(数字).pdf

    Args:
        pdf_filename: PDF文件名（不含路径）

    Returns:
        姓名字符串，如果提取失败则返回None
    """
    # 模式1: 姓名-承诺书.pdf 或 姓名-承诺书(数字).pdf
    pattern1 = r"^(.+?)-承诺书(?:\(\d+\))?\.pdf$"
    match1 = re.match(pattern1, pdf_filename)
    if match1:
        name = match1.group(1).strip()
        # 确保不是空字符串
        if name:
            return name

    # 模式2: 承诺书-姓名.pdf 或 承诺书-姓名(数字).pdf
    pattern2 = r"^承诺书-(.+?)(?:\(\d+\))?\.pdf$"
    match2 = re.match(pattern2, pdf_filename)
    if match2:
        name = match2.group(1).strip()
        # 确保不是空字符串
        if name:
            return name

    # 模式3: 更通用的模式，匹配 "任意内容-任意内容.pdf" 或 "任意内容-任意内容(数字).pdf"
    # 尝试提取第一个部分作为姓名（如果第二个部分包含"承诺书"）
    pattern3 = r"^(.+?)-(.+?)(?:\(\d+\))?\.pdf$"
    match3 = re.match(pattern3, pdf_filename)
    if match3:
        part1 = match3.group(1).strip()
        part2 = match3.group(2).strip()
        # 如果第二个部分包含"承诺书"，则第一个部分是姓名
        if "承诺书" in part2:
            if part1:
                return part1
        # 如果第一个部分包含"承诺书"，则第二个部分是姓名
        elif "承诺书" in part1:
            if part2:
                return part2

    return None


def match_and_copy_pdfs(
    pdf_dir: Path, excel_path: Path, output_dir_name: str = "匹配结果"
):
    """
    匹配PDF文件并移动到目标文件夹

    Args:
        pdf_dir: PDF文件所在文件夹路径
        excel_path: Excel文件路径
        output_dir_name: 输出文件夹名称（默认：匹配结果）
    """
    # 检查PDF文件夹是否存在
    if not pdf_dir.exists():
        print(f"错误: PDF文件夹不存在: {pdf_dir}")
        return

    if not pdf_dir.is_dir():
        print(f"错误: 路径不是文件夹: {pdf_dir}")
        return

    # 读取Excel文件
    print(f"正在读取Excel文件: {excel_path}")
    try:
        df_excel, name_col = read_excel_dataframe(excel_path)
        excel_names = read_excel_names(excel_path)
        print(f"成功读取Excel文件，共 {len(excel_names)} 个姓名")
    except Exception as e:
        print(f"错误: 无法读取Excel文件: {e}")
        return

    # 扫描PDF文件
    pdf_files = list(pdf_dir.glob("*.pdf"))
    print(f"\n找到 {len(pdf_files)} 个PDF文件")

    if len(pdf_files) == 0:
        print("没有找到PDF文件，程序退出")
        return

    # 创建输出文件夹
    output_dir = pdf_dir / output_dir_name
    try:
        output_dir.mkdir(exist_ok=True)
        print(f"输出文件夹: {output_dir}")
    except Exception as e:
        print(f"错误: 无法创建输出文件夹: {e}")
        return

    # 统计信息
    total_count = len(pdf_files)
    matched_count = 0
    unmatched_count = 0
    error_count = 0

    print("\n开始处理PDF文件...")
    print("=" * 80)

    # 处理每个PDF文件
    for pdf_file in pdf_files:
        pdf_filename = pdf_file.name
        name = extract_name_from_filename(pdf_filename)

        if name is None:
            print(f"跳过: {pdf_filename} (无法提取姓名)")
            unmatched_count += 1
            continue

        # 在Excel中查找匹配
        if name not in excel_names:
            print(f"未匹配: {pdf_filename} (姓名: {name} 在Excel中未找到)")
            unmatched_count += 1
            continue

        # 匹配成功，移动文件
        try:
            dest_path = output_dir / pdf_filename
            # 如果目标文件已存在，添加序号
            if dest_path.exists():
                base_name = pdf_file.stem
                ext = pdf_file.suffix
                counter = 1
                while dest_path.exists():
                    new_name = f"{base_name}_{counter}{ext}"
                    dest_path = output_dir / new_name
                    counter += 1
                print(
                    f"移动: {pdf_filename} -> {dest_path.name} (目标文件已存在，已重命名)"
                )
            else:
                print(f"移动: {pdf_filename}")

            shutil.move(str(pdf_file), str(dest_path))
            matched_count += 1
        except Exception as e:
            print(f"错误: 移动失败 {pdf_filename}: {e}")
            error_count += 1

    # 构建PDF文件夹中所有姓名的集合（用于快速查找）
    pdf_names = set()
    for pdf_file in pdf_files:
        name = extract_name_from_filename(pdf_file.name)
        if name:
            pdf_names.add(name)

    # 在Excel中标注匹配状态
    print("\n" + "=" * 80)
    print("正在Excel中标注匹配状态...")
    try:
        # 创建标注列（如果不存在）
        mark_col = "匹配状态"
        if mark_col not in df_excel.columns:
            df_excel[mark_col] = ""

        # 遍历Excel中的每一行，检查是否在PDF文件夹中找到匹配
        success_count = 0
        failed_count = 0

        for idx, row in df_excel.iterrows():
            row_name = str(row[name_col]).strip()
            # 跳过空值
            if pd.isna(row[name_col]) or row_name == "nan" or row_name == "":
                continue

            # 检查该姓名是否在PDF文件夹中找到匹配
            if row_name in pdf_names:
                df_excel.at[idx, mark_col] = "成功"
                success_count += 1
            else:
                df_excel.at[idx, mark_col] = "失败"
                failed_count += 1

        # 保存Excel文件
        # 根据原文件格式选择保存方式
        file_ext = excel_path.suffix.lower()
        if file_ext == ".xls":
            # .xls格式需要使用xlwt，但pandas可能不支持直接写入.xls
            # 尝试保存为.xlsx格式（带备份）
            backup_path = excel_path.with_suffix(".xlsx.backup")
            if backup_path.exists():
                backup_path.unlink()
            excel_path.rename(backup_path)
            new_excel_path = excel_path.with_suffix(".xlsx")
            df_excel.to_excel(new_excel_path, index=False, engine="openpyxl")
            print(f"已保存标注结果到: {new_excel_path}")
            print(f"原文件已备份为: {backup_path}")
        else:
            # .xlsx格式直接保存
            backup_path = excel_path.with_suffix(".backup")
            if backup_path.exists():
                backup_path.unlink()
            shutil.copy2(excel_path, backup_path)
            df_excel.to_excel(excel_path, index=False, engine="openpyxl")
            print(f"已保存标注结果到: {excel_path}")
            print(f"原文件已备份为: {backup_path}")

        print(f"匹配成功: {success_count} 条")
        print(f"匹配失败: {failed_count} 条")
    except Exception as e:
        print(f"警告: 标注Excel文件时出错: {e}")

    # 输出统计信息
    print("\n" + "=" * 80)
    print("处理完成！")
    print(f"总共扫描: {total_count} 个文件")
    print(f"匹配成功: {matched_count} 个")
    print(f"未匹配: {unmatched_count} 个")
    if error_count > 0:
        print(f"处理错误: {error_count} 个")
    print(f"\n匹配的文件已移动到: {output_dir}")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="根据Excel文件中的姓名匹配PDF文件，并将匹配的文件移动到'匹配结果'文件夹"
    )
    parser.add_argument("pdf_dir", type=str, help="PDF文件所在文件夹路径")
    parser.add_argument("excel_path", type=str, help="Excel文件路径（包含姓名列）")
    parser.add_argument(
        "--output-dir",
        type=str,
        default="匹配结果",
        help="输出文件夹名称（默认: 匹配结果）",
    )

    args = parser.parse_args()

    # 转换为Path对象
    pdf_dir = Path(args.pdf_dir).resolve()
    excel_path = Path(args.excel_path).resolve()

    # 执行匹配和复制
    match_and_copy_pdfs(pdf_dir, excel_path, args.output_dir)


if __name__ == "__main__":
    main()
