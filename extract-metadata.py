import re
from typing import Optional, Dict, List
from pathlib import Path
import PyPDF2
import pandas as pd
from difflib import SequenceMatcher
from datetime import datetime

# ===== 配置 =====
JOURNAL_DATABASE_PATH = "2025JCRIMPACTFACTORSDETAILED.xlsx"  # 修改为实际路径
PDF_DIRECTORY = "./papers"                                   # 修改为PDF文件所在目录
OUTPUT_FILE = "parse-results.xlsx"                           # 修改为输出的文件（文件地址+文件名）
TREAT_MODE = "path"                                          # 可选single（单一文件、用于测试）和path（文件夹路径，生产）

# noinspection PyTypeChecker
def load_journal_database(file_path: str = JOURNAL_DATABASE_PATH) -> pd.DataFrame:
    """
    加载期刊影响因子数据库

    Args:
        file_path: xlsx文件路径

    Returns:
        包含期刊信息的DataFrame
    """
    df = pd.read_excel(file_path, usecols=('Journal Name', 'JIF'))  # 按照期刊名和影响因子提取。
    # 不知道为什么，pandas手册上写usecols能接受tuple，但我用tuple就会被IDE骂 (╯▔皿▔)╯
    df.columns = ['journal_name', 'impact_factor']  # 重命名列
    df['journal_name_lower'] = df['journal_name'].str.lower().str.strip()
    return df


def calculate_similarity(str1: str, str2: str) -> float:
    """
    计算两个字符串的相似度

    Args:
        str1: 字符串1
        str2: 字符串2

    Returns:
        相似度分数 (0-1)
    """
    return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()


def get_impact_factor(journal_name: str, df: pd.DataFrame,
                      threshold: float = 0.85) -> Optional[Dict]:
    """
    从本地数据库查询期刊影响因子

    Args:
        journal_name: 期刊名称
        df: 期刊数据库DataFrame
        threshold: 相似度阈值（0-1），默认0.85

    Returns:
        影响因子信息字典或None
    """
    journal_name_clean = journal_name.lower().strip()

    # 精确匹配
    exact_match = df[df['journal_name_lower'] == journal_name_clean]
    if not exact_match.empty:
        row = exact_match.iloc[0]
        return {
            'journal_name': row['journal_name'],
            'impact_factor': row['impact_factor'],
            'match_type': 'exact'
        }

    # 模糊匹配
    df['similarity'] = df['journal_name_lower'].apply(
        lambda x: calculate_similarity(x, journal_name_clean)
    )

    best_match = df.loc[df['similarity'].idxmax()]

    if best_match['similarity'] >= threshold:
        return {
            'journal_name': best_match['journal_name'],
            'impact_factor': best_match['impact_factor'],
            'match_type': 'fuzzy',
            'similarity': round(best_match['similarity'], 3)
        }

    return None


def extract_text_from_pdf(pdf_path: str, max_pages: int = 2) -> str:
    """
    从PDF文件中提取文本（主要检查前几页）

    Args:
        pdf_path: PDF文件路径
        max_pages: 最多读取的页数

    Returns:
        提取的文本内容
    """
    text = ""
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        pages_to_read = min(max_pages, len(pdf_reader.pages))

        for page_num in range(pages_to_read):
            text += pdf_reader.pages[page_num].extract_text()

    return text


def extract_metadata(pdf_path: str) -> Dict[str, str]:
    """
    提取PDF元数据（包括可能的期刊信息）

    Args:
        pdf_path: PDF文件路径

    Returns:
        元数据字典
    """
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        metadata = pdf_reader.metadata or {}

    return {k: v for k, v in metadata.items() if v}


def extract_journal_from_subject(subject: str) -> Optional[str]:
    """
    从subject字段提取期刊名
    通常格式为: [journal name], xxxx, doi:xxx

    Args:
        subject: PDF元数据中的subject字段

    Returns:
        期刊名称或None
    """
    if not subject:
        return None

    # 移除首尾空白
    subject = subject.strip()

    # 方法1: 提取第一个逗号之前的内容
    # 例如: "Nature, 2023, doi:10.1038/xxx" -> "Nature"
    if ',' in subject:
        journal_name = subject.split(',')[0].strip()

        # 验证提取的内容是否像期刊名（包含字母且不全是数字）
        if journal_name and any(c.isalpha() for c in journal_name):
            # 清理可能的前缀符号
            journal_name = re.sub(r'^[^\w\s]+', '', journal_name)
            return journal_name

    # 方法2: 如果没有逗号，尝试提取doi之前的部分
    # 例如: "Nature doi:10.1038/xxx" -> "Nature"
    if 'doi:' in subject.lower():
        parts = re.split(r'\s*doi:', subject, flags=re.IGNORECASE)
        if parts[0].strip():
            return parts[0].strip()

    # 方法3: 如果包含年份，提取年份之前的部分
    # 例如: "Nature 2023" -> "Nature"
    year_match = re.search(r'^(.+?)\s+(?:19|20)\d{2}', subject)
    if year_match:
        return year_match.group(1).strip()

    # 如果都不匹配，返回整个subject（去除数字和特殊字符较多的部分）
    cleaned = re.sub(r'\s*[\d,.:\-]+\s*$', '', subject).strip()
    if cleaned and len(cleaned) > 3:
        return cleaned

    return None

def extract_journal_name(text: str, metadata: Dict[str, str]) -> Optional[str]:
    """
    从文本和元数据中提取期刊名

    Args:
        text: PDF提取的文本
        metadata: PDF元数据

    Returns:
        期刊名称或None
    """
    # 方法1：从元数据的Subject字段提取
    if metadata.get('/Subject'):
        journal_name = extract_journal_from_subject(metadata['/Subject'])
        if journal_name:
            return journal_name

    # 方法2：常见期刊名模式匹配
    patterns = [
        r'Published in:?\s*([A-Z][A-Za-z\s&\-:]+?)(?:\n|,|Vol|\d{4})',
        r'([A-Z][A-Za-z\s&\-:]+?)\s+Vol\.\s*\d+',
        r'Journal:\s*([A-Z][A-Za-z\s&\-:]+)',
        r'©.*?(\b[A-Z][A-Za-z\s&\-:]+?)\s+\d{4}',  # 版权声明中的期刊名
    ]

    for pattern in patterns:
        match = re.search(pattern, text[:2000])
        if match:
            return match.group(1).strip()

    return None


def process_pdf(pdf_path: str, journal_df: pd.DataFrame = None) -> Dict[str, any]:
    """
    主处理函数：读取PDF并获取期刊影响因子

    Args:
        pdf_path: PDF文件路径
        journal_df: 期刊数据库DataFrame（可选，不传则自动加载）

    Returns:
        包含期刊名和影响因子的字典
    """
    # 加载期刊数据库
    if journal_df is None:
        try:
            journal_df = load_journal_database()
        except Exception as e:
            return {
                'status': 'error',
                'message': f'加载期刊数据库失败: {str(e)}'
            }

    # 提取文本和元数据
    try:
        text = extract_text_from_pdf(pdf_path)
        metadata = extract_metadata(pdf_path)
    except Exception as e:
        return {
            'status': 'error',
            'message': f'读取PDF失败: {str(e)}'
        }

    # 提取期刊名
    journal_name = extract_journal_name(text, metadata)

    if not journal_name:
        return {
            'status': 'error',
            'message': '未能识别期刊名称',
            'extracted_text_preview': text[:500]  # 提供前500字符用于调试
        }

    # 查询影响因子
    impact_info = get_impact_factor(journal_name, journal_df)

    if not impact_info:
        return {
            'status': 'not_found',
            'message': '未找到匹配的期刊',
            'extracted_journal_name': journal_name
        }

    return {
        'status': 'success',
        'extracted_journal_name': journal_name,
        'matched_journal_name': impact_info['journal_name'],
        'impact_factor': impact_info['impact_factor'],
        'match_type': impact_info['match_type'],
        'similarity': impact_info.get('similarity')
    }


def find_all_pdfs(directory: str, recursive: bool = True) -> List[Path]:
    """
    查找目录下所有PDF文件

    Args:
        directory: 目录路径
        recursive: 是否递归查找子目录

    Returns:
        PDF文件路径列表
    """
    directory_path = Path(directory)

    if not directory_path.exists():
        raise FileNotFoundError(f"目录不存在: {directory}")

    if recursive:
        pdf_files = list(directory_path.rglob("*.pdf"))
    else:
        pdf_files = list(directory_path.glob("*.pdf"))

    return sorted(pdf_files)


def batch_process_pdfs(pdf_directory: str, journal_df: pd.DataFrame = None,
                       recursive: bool = True) -> List[Dict]:
    """
    批量处理目录下的所有PDF文件

    Args:
        pdf_directory: PDF文件目录
        journal_df: 期刊数据库DataFrame（可选）
        recursive: 是否递归查找子目录

    Returns:
        处理结果列表
    """
    # 查找所有PDF文件
    pdf_files = find_all_pdfs(pdf_directory, recursive)
    total_files = len(pdf_files)

    if total_files == 0:
        print(f"在 {pdf_directory} 中未找到PDF文件")
        return []

    print(f"找到 {total_files} 个PDF文件，开始处理...\n")

    # 预加载数据库
    if journal_df is None:
        journal_df = load_journal_database()

    batch_results = []

    for idx, pdf_file in enumerate(pdf_files, 1):
        print(f"[{idx}/{total_files}] 处理: {pdf_file.name}")

        batch_result = process_pdf(str(pdf_file), journal_df)
        batch_result['file_path'] = str(pdf_file)
        batch_result['file_name'] = pdf_file.name

        if batch_result['status'] == 'success':
            print(f"  ✓ 期刊: {batch_result['matched_journal_name']}")
            print(f"  ✓ IF: {batch_result['impact_factor']}")
        elif batch_result['status'] == 'not_found':
            print(f"  ✗ 未找到: {batch_result.get('extracted_journal_name', 'N/A')}")
        else:
            print(f"  ✗ 错误: {batch_result['message']}")

        batch_results.append(batch_result)
        print()

    return batch_results


def save_results_to_excel(results: List[Dict], output_path: str = None):
    """
    将处理结果保存为Excel文件

    Args:
        results: 处理结果列表
        output_path: 输出文件路径（可选）
    """
    if not results:
        print("没有结果可保存")
        return

    # 准备数据
    data = []
    for save_result in results:
        data.append({
            '文件名': save_result.get('file_name', ''),
            '文件路径': save_result.get('file_path', ''),
            '状态': save_result.get('status', ''),
            '提取的期刊名': save_result.get('extracted_journal_name', ''),
            '匹配的期刊名': save_result.get('matched_journal_name', ''),
            '影响因子': save_result.get('impact_factor', ''),
            '匹配类型': save_result.get('match_type', ''),
            '相似度': save_result.get('similarity', ''),
            '错误信息': save_result.get('message', '')
        })

    df = pd.DataFrame(data)

    # 生成默认文件名
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"journal_impact_factors_{timestamp}.xlsx"

    df.to_excel(output_path, index=False)
    print(f"\n结果已保存到: {output_path}")


def print_summary(results: List[Dict]):
    """
    打印处理结果统计

    Args:
        results: 处理结果列表
    """
    if not results:
        return

    total = len(results)
    success = sum(1 for r in results if r['status'] == 'success')
    not_found = sum(1 for r in results if r['status'] == 'not_found')
    error = sum(1 for r in results if r['status'] == 'error')

    print("=" * 60)
    print("处理统计")
    print("=" * 60)
    print(f"总文件数: {total}")
    print(f"成功匹配: {success} ({success / total * 100:.1f}%)")
    print(f"未找到期刊: {not_found} ({not_found / total * 100:.1f}%)")
    print(f"处理错误: {error} ({error / total * 100:.1f}%)")
    print("=" * 60)


# 主函数
if __name__ == "__main__":
    if TREAT_MODE == "single":
        # ===== 模式1: 处理单个PDF文件 =====
        print("模式1: 单个文件处理")
        print("-" * 60)

        single_pdf = "your_paper.pdf"
        result = process_pdf(single_pdf)

        if result['status'] == 'success':
            print(f"提取的期刊名: {result['extracted_journal_name']}")
            print(f"匹配的期刊名: {result['matched_journal_name']}")
            print(f"影响因子: {result['impact_factor']}")
            print(f"匹配类型: {result['match_type']}")
            if result.get('similarity'):
                print(f"相似度: {result['similarity']}")
        else:
            print(f"状态: {result['status']}")
            print(f"信息: {result['message']}")

        print("\n" + "=" * 60 + "\n")


    # ===== 模式2: 批量处理指定目录下所有PDF =====
    print("模式2: 批量处理目录")
    print("-" * 60)

    # 方式A: 使用默认配置的目录
    results = batch_process_pdfs(PDF_DIRECTORY, recursive=True)

    # 打印统计信息
    print_summary(results)

    # 使用指定的输出文件名
    save_results_to_excel(results, OUTPUT_FILE)