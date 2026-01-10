#!/usr/bin/env python3
"""
PDF文件匹配复制脚本
根据Excel文件中的身份证号匹配PDF文件，并将匹配的文件复制到"匹配结果"文件夹
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
    读取Excel文件，返回DataFrame和身份证号列名

    Args:
        excel_path: Excel文件路径

    Returns:
        (DataFrame, 身份证号列名)元组

    Raises:
        FileNotFoundError: Excel文件不存在
        ValueError: Excel文件格式错误或找不到身份证号列
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

        # 查找"身份证号"列
        id_col = None
        for col in df.columns:
            col_str = str(col).strip()
            if "身份证号" in col_str or "身份证" in col_str:
                id_col = col
                break

        if id_col is None:
            raise ValueError(
                f"在Excel文件中找不到'身份证号'列。可用列: {list(df.columns)}"
            )

        return (df, id_col)

    except Exception as e:
        if isinstance(e, (FileNotFoundError, ValueError)):
            raise
        raise ValueError(f"读取Excel文件时出错: {str(e)}")


def read_excel_id_numbers(excel_path: Path) -> Set[str]:
    """
    读取Excel文件，提取身份证号列的所有值

    Args:
        excel_path: Excel文件路径

    Returns:
        身份证号集合

    Raises:
        FileNotFoundError: Excel文件不存在
        ValueError: Excel文件格式错误或找不到身份证号列
    """
    """读取Excel文件中的身份证号集合（兼容旧接口）"""
    df, id_col = read_excel_dataframe(excel_path)

    # 提取所有身份证号，去除空值和NaN
    id_numbers = set()
    for idx, row in df.iterrows():
        id_value = str(row[id_col]).strip()
        # 跳过空值和NaN
        if pd.isna(row[id_col]) or id_value == "nan" or id_value == "":
            continue
        # 处理可能的科学计数法格式
        if "." in id_value:
            try:
                id_value = str(int(float(id_value)))
            except (ValueError, OverflowError):
                pass
        id_numbers.add(id_value)

    return id_numbers


def extract_name_and_id_from_filename(pdf_filename: str) -> Optional[tuple[str, str]]:
    """
    从PDF文件名中提取姓名和身份证号

    文件名格式：协商解除劳动合同协议书_姓名身份证号.pdf
    身份证号格式：18位，最后一位可能是0-9或X

    Args:
        pdf_filename: PDF文件名（不含路径）

    Returns:
        (姓名, 身份证号)元组，如果提取失败则返回None
    """
    # 移除扩展名
    name_without_ext = pdf_filename.replace(".pdf", "")

    # 使用正则表达式匹配固定格式
    # 身份证号：17位数字 + 1位数字或X
    pattern = r"协商解除劳动合同协议书_(.+?)(\d{17}[\dX])\.pdf"
    match = re.search(pattern, pdf_filename, re.IGNORECASE)
    if match:
        name = match.group(1)
        id_number = match.group(2).upper()  # 统一转换为大写X
        return (name, id_number)

    # 备用方法：从文件名末尾提取最后18位（17位数字+1位数字或X）
    # 查找匹配身份证号格式的字符串（17位数字+1位数字或X）
    id_pattern = r"\d{17}[\dX]"
    id_matches = re.findall(id_pattern, name_without_ext, re.IGNORECASE)
    if id_matches:
        # 取最后一个匹配（最接近文件末尾的）
        last_id = id_matches[-1].upper()
        # 提取下划线后的内容，去掉最后的身份证号
        prefix = "协商解除劳动合同协议书_"
        if name_without_ext.startswith(prefix):
            remaining = name_without_ext[len(prefix) :]
            if remaining.upper().endswith(last_id):
                name = remaining[:-18]
                return (name, last_id)

    return None


def extract_id_number_from_filename(pdf_filename: str) -> Optional[str]:
    """
    从PDF文件名中提取身份证号（固定18位）

    文件名格式：协商解除劳动合同协议书_姓名身份证号.pdf
    身份证号格式：18位，最后一位可能是0-9或X
    提取逻辑：从文件名末尾提取最后18位（17位数字+1位数字或X）

    Args:
        pdf_filename: PDF文件名（不含路径）

    Returns:
        身份证号字符串（18位，最后一位可能是X），如果提取失败则返回None
    """
    result = extract_name_and_id_from_filename(pdf_filename)
    if result:
        return result[1]
    return None


def match_and_copy_pdfs(
    pdf_dir: Path, excel_path: Path, output_dir_name: str = "匹配结果"
):
    """
    匹配PDF文件并复制到目标文件夹

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
        df_excel, id_col = read_excel_dataframe(excel_path)
        excel_id_numbers = read_excel_id_numbers(excel_path)
        print(f"成功读取Excel文件，共 {len(excel_id_numbers)} 个身份证号")
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
        name_id_result = extract_name_and_id_from_filename(pdf_filename)

        if name_id_result is None:
            print(f"跳过: {pdf_filename} (无法提取姓名和身份证号)")
            unmatched_count += 1
            continue

        name, id_number = name_id_result

        # 在Excel中查找匹配
        if id_number not in excel_id_numbers:
            print(
                f"未匹配: {pdf_filename} (姓名: {name}, 身份证号: {id_number} 在Excel中未找到)"
            )
            unmatched_count += 1
            continue

        # 匹配成功，复制文件
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
                    f"复制: {pdf_filename} -> {dest_path.name} (目标文件已存在，已重命名)"
                )
            else:
                print(f"复制: {pdf_filename}")

            shutil.copy2(pdf_file, dest_path)
            matched_count += 1
        except Exception as e:
            print(f"错误: 复制失败 {pdf_filename}: {e}")
            error_count += 1

    # 构建PDF文件夹中所有身份证号的集合（用于快速查找）
    # 统一转换为大写，确保X能正确匹配
    pdf_id_numbers = set()
    for pdf_file in pdf_files:
        name_id_result = extract_name_and_id_from_filename(pdf_file.name)
        if name_id_result:
            _, id_number = name_id_result
            pdf_id_numbers.add(id_number.upper())  # 统一转换为大写

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
            row_id = str(row[id_col]).strip()
            # 跳过空值
            if pd.isna(row[id_col]) or row_id == "nan" or row_id == "":
                continue

            # 处理可能的科学计数法格式
            if "." in row_id:
                try:
                    row_id = str(int(float(row_id)))
                except (ValueError, OverflowError):
                    pass

            # 统一转换为大写，确保X能正确匹配
            row_id = row_id.upper()

            # 检查该身份证号是否在PDF文件夹中找到匹配
            if row_id in pdf_id_numbers:
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
    print(f"\n匹配的文件已复制到: {output_dir}")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="根据Excel文件中的身份证号匹配PDF文件，并将匹配的文件复制到'匹配结果'文件夹"
    )
    parser.add_argument("pdf_dir", type=str, help="PDF文件所在文件夹路径")
    parser.add_argument("excel_path", type=str, help="Excel文件路径（包含身份证号列）")
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
