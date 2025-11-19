"""
主处理脚本：处理Excel数据并生成图表

功能：
1. 调用 process_excel_data.py 处理原始 Excel 文件，生成标准化文件
2. 调用 auto_chart_generator.py 生成图表

输入：报告数据.xlsx（或指定的Excel文件）
输出：
  - 报告数据_标准化.xlsx（标准化数据文件）
  - charts/generated/（图表输出目录）
"""

import sys
from pathlib import Path
from typing import Optional

# 导入处理模块
from process_excel_data import process_excel_file
from auto_chart_generator import ChartGenerator, DEFAULT_LOGIC_PATH, OUTPUT_DIR


def process_and_generate(
    input_excel: Path,
    output_standardized: Optional[Path] = None,
    logic_path: Optional[Path] = None,
    output_dir: Optional[Path] = None,
) -> bool:
    """
    处理Excel文件并生成图表
    
    Args:
        input_excel: 输入的原始Excel文件路径（如：报告数据.xlsx）
        output_standardized: 输出的标准化Excel文件路径（默认：报告数据_标准化.xlsx）
        logic_path: 图表逻辑定义文件路径（默认：chart_generation_logic.jsonl）
        output_dir: 图表输出目录（默认：charts/generated）
    
    Returns:
        bool: 处理是否成功
    """
    # 设置默认路径
    if output_standardized is None:
        # 如果输入是"报告数据.xlsx"，输出"报告数据_标准化.xlsx"
        if input_excel.stem == "报告数据":
            output_standardized = input_excel.parent / "报告数据_标准化.xlsx"
        else:
            # 否则在输入文件名后添加"_标准化"
            output_standardized = input_excel.parent / f"{input_excel.stem}_标准化.xlsx"
    
    if logic_path is None:
        logic_path = DEFAULT_LOGIC_PATH
    
    if output_dir is None:
        output_dir = OUTPUT_DIR
    
    # 步骤1：标准化Excel文件
    print("\n" + "="*80)
    print("步骤 1/2: 标准化Excel文件")
    print("="*80)
    
    success = process_excel_file(input_excel, output_standardized)
    
    if not success:
        print("\n✗ Excel标准化失败，终止处理")
        return False
    
    # 步骤2：生成图表
    print("\n" + "="*80)
    print("步骤 2/2: 生成图表")
    print("="*80)
    
    try:
        print(f"\n使用标准化文件: {output_standardized}")
        print(f"使用逻辑文件: {logic_path}")
        print(f"输出目录: {output_dir}\n")
        
        generator = ChartGenerator(
            logic_path=logic_path,
            excel_path=output_standardized,
            output_dir=output_dir,
        )
        generator.generate_all()
        
        print("\n" + "="*80)
        print("✓ 所有处理完成！")
        print("="*80)
        print(f"\n输出文件：")
        print(f"  - 标准化数据: {output_standardized}")
        print(f"  - 图表目录: {output_dir}")
        print()
        
        return True
        
    except Exception as e:
        print(f"\n✗ 图表生成失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """主函数，支持命令行参数"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="处理Excel数据并生成图表",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  # 使用默认文件（报告数据.xlsx）
  python process_main.py
  
  # 指定输入文件
  python process_main.py -i "A项目资料补充_251111/附件7_报告数据分析_250625R1.xlsx"
  
  # 指定所有参数
  python process_main.py -i "报告数据.xlsx" -o "报告数据_标准化.xlsx" -l "chart_generation_logic.jsonl"
        """
    )
    
    parser.add_argument(
        "-i", "--input",
        type=Path,
        default=Path("附件7_报告数据分析_250625R1.xlsx"),
        help="输入的原始Excel文件路径（默认：报告数据.xlsx）"
    )
    
    parser.add_argument(
        "-o", "--output",
        type=Path,
        default=None,
        help="输出的标准化Excel文件路径（默认：根据输入文件名自动生成）"
    )
    
    parser.add_argument(
        "-l", "--logic",
        type=Path,
        default=None,
        help=f"图表逻辑定义文件路径（默认：{DEFAULT_LOGIC_PATH}）"
    )
    
    parser.add_argument(
        "-d", "--output-dir",
        type=Path,
        default=None,
        help=f"图表输出目录（默认：{OUTPUT_DIR}）"
    )
    
    args = parser.parse_args()
    
    # 检查输入文件是否存在
    if not args.input.exists():
        print(f"✗ 错误：输入文件不存在: {args.input}")
        print(f"  请检查文件路径是否正确")
        sys.exit(1)
    
    # 执行处理
    success = process_and_generate(
        input_excel=args.input,
        output_standardized=args.output,
        logic_path=args.logic,
        output_dir=args.output_dir,
    )
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()

