#!/usr/bin/env python3
"""
PDF文件重命名脚本
根据Excel文件中的合同编号匹配信息，将PDF文件从"合同编号"格式重命名为"姓名+身份证号"格式
"""

import re
import argparse
from pathlib import Path
from typing import Dict, Tuple, Optional
import pandas as pd


def read_excel_mapping(excel_path: Path) -> Dict[str, Tuple[str, str]]:
    """
    读取Excel文件，构建合同编号到(姓名, 身份证号)的映射字典
    
    Args:
        excel_path: Excel文件路径
        
    Returns:
        字典，键为合同编号（字符串），值为(姓名, 身份证号)的元组
        
    Raises:
        FileNotFoundError: Excel文件不存在
        ValueError: Excel文件格式错误或找不到必要的列
    """
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel文件不存在: {excel_path}")
    
    try:
        # 根据文件扩展名选择首选引擎
        file_ext = excel_path.suffix.lower()
        if file_ext == '.xls':
            preferred_engines = ['xlrd', 'openpyxl']  # .xls扩展名，但可能是xlsx格式
        elif file_ext in ['.xlsx', '.xlsm']:
            preferred_engines = ['openpyxl', 'xlrd']
        else:
            # 扩展名不明确，尝试两种引擎
            preferred_engines = ['xlrd', 'openpyxl']
        
        # 尝试使用引擎读取文件
        df = None
        last_error = None
        for engine in preferred_engines:
            try:
                df = pd.read_excel(excel_path, engine=engine)
                break  # 成功读取，退出循环
            except Exception as e:
                last_error = e
                continue  # 尝试下一个引擎
        
        if df is None:
            raise ValueError(f"无法读取Excel文件: {last_error}")
        
        # 查找"合同编号"列
        contract_col = None
        name_col = None
        id_col = None
        
        for col in df.columns:
            col_str = str(col).strip()
            if '合同编号' in col_str:
                contract_col = col
            elif '姓名' in col_str:
                name_col = col
            elif '身份证号' in col_str or '身份证' in col_str:
                id_col = col
        
        if contract_col is None:
            raise ValueError(f"在Excel文件中找不到'合同编号'列。可用列: {list(df.columns)}")
        if name_col is None:
            raise ValueError(f"在Excel文件中找不到'姓名'列。可用列: {list(df.columns)}")
        if id_col is None:
            raise ValueError(f"在Excel文件中找不到'身份证号'列。可用列: {list(df.columns)}")
        
        # 构建映射字典
        mapping = {}
        for idx, row in df.iterrows():
            contract_num = str(row[contract_col]).strip()
            name = str(row[name_col]).strip()
            id_num = str(row[id_col]).strip()
            
            # 跳过空值
            if pd.isna(row[contract_col]) or contract_num == 'nan' or contract_num == '':
                continue
            
            # 将合同编号转换为字符串（去除可能的科学计数法格式）
            if '.' in contract_num:
                try:
                    contract_num = str(int(float(contract_num)))
                except (ValueError, OverflowError):
                    pass
            
            mapping[contract_num] = (name, id_num)
        
        return mapping
    
    except Exception as e:
        if isinstance(e, (FileNotFoundError, ValueError)):
            raise
        raise ValueError(f"读取Excel文件时出错: {str(e)}")


def extract_contract_number(pdf_filename: str) -> Optional[str]:
    """
    从PDF文件名中提取合同编号
    
    Args:
        pdf_filename: PDF文件名（不含路径）
        
    Returns:
        合同编号字符串，如果格式不匹配则返回None
    """
    pattern = r'协商解除劳动合同协议书_(\d+)\.pdf'
    match = re.match(pattern, pdf_filename)
    if match:
        return match.group(1)
    return None


def rename_pdf_files(target_dir: Path, excel_filename: str = "协商解除函签署名单-608人.xls"):
    """
    批量重命名PDF文件
    
    Args:
        target_dir: 目标文件夹路径
        excel_filename: Excel文件名（固定名称）
    """
    # 检查目标文件夹是否存在
    if not target_dir.exists():
        print(f"错误: 目标文件夹不存在: {target_dir}")
        return
    
    if not target_dir.is_dir():
        print(f"错误: 路径不是文件夹: {target_dir}")
        return
    
    # Excel文件路径
    excel_path = target_dir / excel_filename
    
    # 读取Excel映射
    print(f"正在读取Excel文件: {excel_path}")
    try:
        mapping = read_excel_mapping(excel_path)
        print(f"成功读取Excel文件，共 {len(mapping)} 条记录")
    except Exception as e:
        print(f"错误: 无法读取Excel文件: {e}")
        return
    
    # 扫描PDF文件
    pdf_files = list(target_dir.glob("*.pdf"))
    print(f"\n找到 {len(pdf_files)} 个PDF文件")
    
    if len(pdf_files) == 0:
        print("没有找到PDF文件，程序退出")
        return
    
    # 统计信息
    success_count = 0
    failed_count = 0
    skipped_count = 0
    
    print("\n开始处理PDF文件...")
    print("=" * 80)
    
    # 处理每个PDF文件
    for pdf_file in pdf_files:
        pdf_filename = pdf_file.name
        contract_num = extract_contract_number(pdf_filename)
        
        if contract_num is None:
            print(f"跳过: {pdf_filename} (文件名格式不匹配)")
            skipped_count += 1
            continue
        
        # 在映射中查找
        if contract_num not in mapping:
            print(f"未匹配: {pdf_filename} (合同编号 {contract_num} 在Excel中未找到)")
            failed_count += 1
            continue
        
        name, id_num = mapping[contract_num]
        
        # 检查姓名和身份证号是否有效
        if pd.isna(name) or str(name).strip() == '' or str(name).strip() == 'nan':
            print(f"警告: {pdf_filename} (合同编号 {contract_num} 对应的姓名为空)")
            failed_count += 1
            continue
        
        if pd.isna(id_num) or str(id_num).strip() == '' or str(id_num).strip() == 'nan':
            print(f"警告: {pdf_filename} (合同编号 {contract_num} 对应的身份证号为空)")
            failed_count += 1
            continue
        
        # 构建新文件名
        new_filename = f"协商解除劳动合同协议书_{name}{id_num}.pdf"
        new_path = target_dir / new_filename
        
        # 检查目标文件是否已存在
        if new_path.exists() and new_path != pdf_file:
            print(f"跳过: {pdf_filename} -> {new_filename} (目标文件已存在)")
            failed_count += 1
            continue
        
        # 执行重命名
        try:
            pdf_file.rename(new_path)
            print(f"成功: {pdf_filename} -> {new_filename}")
            success_count += 1
        except Exception as e:
            print(f"错误: 重命名失败 {pdf_filename}: {e}")
            failed_count += 1
    
    # 输出统计信息
    print("=" * 80)
    print(f"\n处理完成！")
    print(f"总计: {len(pdf_files)} 个文件")
    print(f"成功重命名: {success_count} 个")
    print(f"未匹配/失败: {failed_count} 个")
    print(f"格式不匹配: {skipped_count} 个")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="根据Excel文件中的合同编号信息批量重命名PDF文件"
    )
    parser.add_argument(
        "target_dir",
        type=str,
        help="目标文件夹路径（包含需要处理的PDF文件和Excel文件）"
    )
    parser.add_argument(
        "--excel",
        type=str,
        default="协商解除函签署名单-608人.xls",
        help="Excel文件名（默认: 协商解除函签署名单-608人.xls）"
    )
    
    args = parser.parse_args()
    
    # 转换为Path对象
    target_dir = Path(args.target_dir).resolve()
    
    # 执行重命名
    rename_pdf_files(target_dir, args.excel)


if __name__ == "__main__":
    main()
