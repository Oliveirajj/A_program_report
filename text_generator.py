"""
文本生成模块：调用qwen3-vl-plus模型生成分析文本
"""
from __future__ import annotations

import base64
import os
import json
import re
from pathlib import Path
import string
from typing import Dict, List, Optional

try:
    from dotenv import load_dotenv
    import dashscope
    from dashscope import MultiModalConversation
except ImportError as e:
    raise ImportError(
        f"缺少必要的依赖库: {e}\n"
        "请运行: pip install dashscope python-dotenv"
    ) from e

# 加载.env文件
load_dotenv()

# 从环境变量读取API密钥
QWEN_API_KEY = os.getenv("QWEN_API_KEY")
if not QWEN_API_KEY:
    raise ValueError(
        "未找到QWEN_API_KEY环境变量。请在.env文件中设置QWEN_API_KEY=your_api_key"
    )

# 设置DashScope API密钥
dashscope.api_key = QWEN_API_KEY

TEMPLATE_SECTIONS_PATH = Path("template_sections.json")
_TEMPLATE_SECTIONS_CACHE: Optional[Dict[str, Dict[str, str]]] = None


def encode_image_to_base64(image_path: Path) -> str:
    """将图片文件编码为base64字符串"""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")


def _build_reference_label(index: int) -> str:
    """将序号转换为参考图标签（参考图A, 参考图B, ...）"""
    letters = string.ascii_uppercase
    result = ""
    i = index
    while i > 0:
        i, rem = divmod(i - 1, 26)
        result = letters[rem] + result
    return f"参考图{result}"


def load_reference_images(ref_dir: Path) -> List[dict]:
    """
    从ref文件夹加载所有参考图片，返回base64编码的图片列表
    
    Args:
        ref_dir: 参考图片文件夹路径
        
    Returns:
        包含图片信息的字典列表，格式：[{"image": base64_string, "name": filename}, ...]
    """
    if not ref_dir.exists():
        return []
    
    images = []
    processed_paths = set()  # 使用集合跟踪已处理的文件路径，避免重复加载
    
    # 统一收集并排序图片，保证编号稳定
    candidate_paths: List[Path] = []
    for ext in ["*.png", "*.jpg", "*.jpeg"]:
        candidate_paths.extend(sorted(ref_dir.glob(ext), key=lambda p: p.name.lower()))
    
    for idx, image_path in enumerate(candidate_paths, start=1):
        abs_path = image_path.resolve()
        if abs_path in processed_paths:
            continue
        processed_paths.add(abs_path)
        try:
            base64_image = encode_image_to_base64(image_path)
            images.append({
                "image": base64_image,
                "name": image_path.name,
                "path": str(abs_path),
                "label": _build_reference_label(idx),
                "stem": image_path.stem,
            })
        except Exception as e:
            print(f"警告：无法加载图片 {image_path}: {e}")
    
    return images


def load_template_sections() -> Dict[str, Dict[str, str]]:
    """懒加载章节模板内容"""
    global _TEMPLATE_SECTIONS_CACHE
    if _TEMPLATE_SECTIONS_CACHE is None:
        if TEMPLATE_SECTIONS_PATH.exists():
            with TEMPLATE_SECTIONS_PATH.open("r", encoding="utf-8") as f:
                _TEMPLATE_SECTIONS_CACHE = json.load(f)
        else:
            _TEMPLATE_SECTIONS_CACHE = {}
    return _TEMPLATE_SECTIONS_CACHE


def extract_template_key(chapter_title: str) -> Optional[str]:
    """从章节标题中提取形如3.1.1的编号"""
    match = re.search(r"\d+\.\d+\.\d+", chapter_title)
    if match:
        return match.group(0)
    return None


def get_template_guidance(chapter_title: str) -> Optional[Dict[str, object]]:
    """获取模板章节的文本与字数参考"""
    template_sections = load_template_sections()
    section_key = extract_template_key(chapter_title)
    if not section_key:
        return None
    template_entry = template_sections.get(section_key)
    if not template_entry:
        return None

    template_text = template_entry.get("text", "").strip()
    flattened = template_text.replace("\n", "")
    char_length = len(flattened) or len(template_text)
    lower_bound = max(150, int(char_length * 0.8))
    upper_bound = max(lower_bound + 60, int(char_length * 1.2))

    return {
        "key": section_key,
        "title": template_entry.get("title", section_key),
        "text": template_text,
        "length_range": (lower_bound, upper_bound),
        "char_length": char_length,
    }


def build_prompt(
    chapter_title: str,
    chapter_content: str,
    reference_images: List[dict],
    chapter_charts: Optional[List[dict]] = None,
) -> List[dict]:
    """
    构建发送给模型的提示词
    
    Args:
        chapter_title: 章节标题
        chapter_content: 章节内容（段落文本和图表描述）
        reference_images: 参考图片列表
        
    Returns:
        消息列表，格式符合DashScope API要求
    """
    chapter_charts = chapter_charts or []
    template_guidance = get_template_guidance(chapter_title)
    if template_guidance:
        lower, upper = template_guidance["length_range"]  # type: ignore[index]
        guidance_block = (
            f"模板章节：{template_guidance['title']}\n"
            f"字数参考：约{lower}-{upper}字\n"
            f"模板示例文本：\n{template_guidance['text']}\n"
        )
        length_requirement = f"5. 生成字数需参考模板，控制在约{lower}-{upper}字之间，结构清晰"
    else:
        guidance_block = "模板章节：未找到对应模板，仅保持专业能耗分析风格\n"
        length_requirement = "5. 字数约200-500字，可根据内容灵活调整，保持结构完整"

    if reference_images:
        reference_lines = "\n".join(
            f"{img['label']}（文件名：{img['name']}）" for img in reference_images
        )
        reference_block = (
            f"参考图片列表：\n{reference_lines}\n"
            "仅在确实使用了对应图片信息时引用其编号。\n"
        )
    else:
        reference_block = "参考图片：当前章节无额外参考图片。\n"

    if chapter_charts:
        chart_lines = []
        for chart in chapter_charts:
            chart_id = chart.get("chart_id") or ""
            chart_name = chart.get("chart_name") or ""
            if chart_name:
                chart_lines.append(f"{chart_id}（{chart_name}）")
            else:
                chart_lines.append(chart_id)
        chart_lines_text = "\n".join(chart_lines)
        chart_block = (
            f"章节图表列表：\n{chart_lines_text}\n"
            "当引用这些内部图表数据、指标或结论时，请用对应的CH编号进行标记。\n"
        )
    else:
        chart_block = "章节图表列表：当前章节暂未配置自动生成图表或无需引用。\n"

    if reference_images or chapter_charts:
        marker_requirement = (
            "6. 当引用上述参考图片或章节图表信息时，请在相关句子或段落末尾添加括号标记，"
            "标签必须是“参考图X”或“CHxxx”格式。\n"
            "7. 若同一句引用多条资料，使用“|”在括号内分隔，例如(参考图A|CH004)。\n"
            "8. 未引用任何资料时不要添加标记，且只输出分析文本，不要包含标题或其他格式标记"
        )
    else:
        marker_requirement = "6. 只输出分析文本，不要包含标题或其他格式标记"
    
    # 构建文本提示词
    text_prompt = f"""你是一位专业的能耗分析报告撰写专家。请根据附件模板所示的章节文字风格与字数要求，生成与当前章节相匹配的综合分析文本。

章节标题：{chapter_title}

模板参考信息：
{guidance_block}

章节内容：
{chapter_content}

要求：
1. 文字风格参考专业能耗分析报告，使用客观、准确、专业的表述
2. 分析要基于章节中的数据和图表信息
3. 语言简洁明了，逻辑清晰
4. 如果章节内容较少，可以适当补充合理的分析观点
{length_requirement}
{marker_requirement}

参考图片说明：
以下提供了{len(reference_images)}张参考图片，这些图片可能包含相关的行业标准、对比数据或参考案例。你可以参考这些图片中的信息来辅助分析，但主要分析应基于章节内容。如果参考图片与当前章节不相关，可以忽略。
{reference_block}

章节图表说明：
下方列出了本章节可引用的自动生成图表（CH编号），当你在分析中引用这些图表的数据或结论时，请添加相应的CH编号标记。
{chart_block}

请生成分析文本："""

    # 构建消息内容
    content = [{"text": text_prompt}]
    
    # 添加所有参考图片
    # DashScope API支持直接使用base64字符串，或使用file://协议
    for img_info in reference_images:
        # 根据图片文件名判断MIME类型
        img_name = img_info['name'].lower()
        if img_name.endswith('.png'):
            mime_type = "image/png"
        elif img_name.endswith(('.jpg', '.jpeg')):
            mime_type = "image/jpeg"
        else:
            mime_type = "image/png"  # 默认使用PNG
        
        content.append({
            "image": f"data:{mime_type};base64,{img_info['image']}"
        })
    
    messages = [
        {
            "role": "user",
            "content": content
        }
    ]
    
    return messages


def generate_analysis_text(
    chapter_title: str,
    chapter_content: str,
    reference_images: List[dict],
    chapter_charts: Optional[List[dict]] = None,
    model: str = "qwen-vl-plus"
) -> str:
    """
    调用qwen3-vl-plus模型生成分析文本
    
    Args:
        chapter_title: 章节标题
        chapter_content: 章节内容
        reference_images: 参考图片列表
        chapter_charts: 当前章节可引用的图表元数据列表
        model: 模型名称，默认为qwen-vl-plus
        
    Returns:
        生成的分析文本
    """
    messages = build_prompt(
        chapter_title,
        chapter_content,
        reference_images,
        chapter_charts=chapter_charts,
    )
    
    try:
        response = MultiModalConversation.call(
            model=model,
            messages=messages,
            max_tokens=2000
        )
        
        if response.status_code == 200:
            # 提取生成的文本
            # DashScope API返回格式：response.output.choices[0].message.content
            content = response.output.choices[0].message.content
            # content可能是字符串或包含text字段的字典
            if isinstance(content, str):
                output_text = content
            elif isinstance(content, list) and len(content) > 0:
                # 如果是列表，取第一个元素的text字段
                if isinstance(content[0], dict) and "text" in content[0]:
                    output_text = content[0]["text"]
                else:
                    output_text = str(content[0])
            elif isinstance(content, dict) and "text" in content:
                output_text = content["text"]
            else:
                output_text = str(content)
            
            return output_text.strip()
        else:
            error_msg = f"API调用失败: {response.message if hasattr(response, 'message') else '未知错误'}"
            print(f"错误: {error_msg}")
            return f"[文本生成失败: {error_msg}]"
            
    except Exception as e:
        error_msg = f"生成文本时发生错误: {str(e)}"
        print(f"错误: {error_msg}")
        import traceback
        traceback.print_exc()
        return f"[文本生成失败: {error_msg}]"


if __name__ == "__main__":
    # 测试代码
    ref_dir = Path("ref")
    images = load_reference_images(ref_dir)
    print(f"加载了 {len(images)} 张参考图片")
    
    test_title = "3.1 总能耗及密度分析"
    test_content = "本章节包含建筑总能耗数据表和碳排放信息表。"
    
    result = generate_analysis_text(test_title, test_content, images)
    print("\n生成的文本：")
    print(result)

