#! /usr/bin/env python
# @Time    : 2026/1/11 19:51
# @Author  : afish
# @File    : set_style.py
from __future__ import annotations

from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from wordformat.pipeline import PipelineStage

from wordformat.log_config import logger
from wordformat.pipeline.context import FormatContext
from wordformat.pipeline.stages import (
    DocumentSavingStage,
    FormattingExecutionStage,
    LoadConfigStage,
    LoadDocxStage,
    ParagraphAlignmentStage,
    StyleDefinitionFixStage,
    SummaryGenerationStage,
    TreeBuildingStage,
    TreeNormalizationStage,
)
from wordformat.pipeline.stages_md import (
    DocumentCreationStage,
    LoadMarkdownStage,
    MarkdownParseStage,
)


def auto_format_thesis_document(
    jsonpath: str | list,
    docxpath: str,
    configpath: Optional[str] = None,
    savepath: str = "output/",
    check=True,
    presets: Optional[list] = None,
):
    """自动对学位论文文档进行格式校验与批注。

    该函数根据结构化 JSON 描述和 YAML 格式配置，对指定的 Word 文档进行格式合规性检查，
    并在不符合规范的位置插入批注（comments）。主要用于学术论文（如本科/硕士/博士论文）
    的自动化格式审查。

    流程说明：
        1. 从 JSON 文件加载文档逻辑结构树；
        2. 加载 Word 文档，并将每个非空段落匹配到对应的结构节点；
        4. 对特定子树（如中英文摘要、参考文献）执行节点提升操作，确保内容节点正确挂载；
        5. 遍历所有结构节点，依据配置文件中的格式规则进行校验，并在文档中添加批注；
        6. 保存带批注的文档到指定路径。

    Args:
        check (bool): 用来控制是仅检查还是仅修改
        jsonpath (str): 文档逻辑结构的 JSON 文件路径 或 json 数据，描述各章节/段落的语义类型。
        docxpath (str): 待处理的原始 Word (.docx) 文档路径。
        savepath (str): 处理完成后带批注的文档保存路径。
        configpath (Optional[str]): 格式规范配置文件（YAML）路径，支持继承与合并。
                                 为 None 时使用内置默认配置。
        presets (Optional[list]): 预设名列表（presets/ 目录下的 .py 插件），按顺序加载，
                                 后加载的预设覆盖先加载的。

    Side Effects:
        - 读取 jsonpath、docxpath 和 configpath 指定的文件；
        - 在 docx 文档中插入批注（不修改原文内容，仅添加审阅意见）；
        - 将结果文档写入 savepath。

    Example:
        >>> auto_format_thesis_document(
        ...     docxpath="draft.docx",
        ...     jsonpath="thesis_structure.json",
        ...     configpath="format_rules.yaml",
        ...     savepath="output/",
        ...     check=True,
        ...     presets=["tsinghua"]
        ... )
    """

    ctx = FormatContext(
        docx_path=docxpath,
        json_path=jsonpath,
        config_path=configpath,
        save_dir=savepath,
        check=check,
    )
    _prepare_hooks(ctx, presets)
    # 2. 组装流水线
    pipeline: list[PipelineStage] = [
        LoadConfigStage(),
        LoadDocxStage(),
        TreeBuildingStage(),
        ParagraphAlignmentStage(),
        TreeNormalizationStage(),
        StyleDefinitionFixStage(),
        FormattingExecutionStage(),
        SummaryGenerationStage(),
        DocumentSavingStage(),
    ]
    for stage in pipeline:
        ctx = stage.process(ctx)
    return ctx.output_path


def _prepare_hooks(ctx, presets) -> None:
    """注册内置机制 handlers，并装配默认文档领域（thesis），再加载预设。

    顺序：内置机制 → 文档领域（可覆盖/补充内置逻辑） → 预设（最后注册，
    可 unregister 禁用前述任何回调）。
    """
    from wordformat.domains import load_domain
    from wordformat.handlers import register_builtin_handlers

    register_builtin_handlers()
    load_domain("thesis")
    _load_presets_into_context(ctx, presets)


def _load_presets_into_context(ctx, presets) -> None:
    """加载预设列表：注册 hook 回调，manifest 存入 ctx 供配置合并。"""
    if not presets:
        return
    from wordformat.preset_loader import load_presets

    ctx.preset_manifests = load_presets(list(presets))
    logger.info(f"已加载预设: {', '.join(presets)}")


def md_to_docx(
    md_path: str,
    config_path: str | None = None,
    save_dir: str = "output/",
    presets: Optional[list] = None,
):
    """将 Markdown 文件转换为格式化后的 .docx 文档。

    流程：
        1. 加载 YAML 格式配置（可选）；
        2. 读取并解析 Markdown 文件为扁平段落列表；
        3. 构建 FormatNode 文档树；
        4. 创建新的 Word 文档，逐节点生成段落并填入文本；
        5. 修正样式定义；
        6. 应用格式规则；
        7. 后处理（标题编号、引用超链接）；
        8. 保存输出 .docx 文件。

    Args:
        md_path: Markdown 源文件路径。
        config_path: YAML 格式规范配置文件路径，为 None 时使用内置默认配置。
        save_dir: 输出目录。
        presets: 预设名列表，按顺序加载（后加载覆盖先加载）。

    Returns:
        生成的 .docx 文件路径。
    """
    from pathlib import Path

    from wordformat.utils import ensure_directory_exists, get_file_name

    ctx = FormatContext(
        md_path=md_path,
        config_path=config_path or "",
        save_dir=save_dir,
        check=False,
    )
    _prepare_hooks(ctx, presets)

    pipeline: list[PipelineStage] = [
        LoadConfigStage(),
        LoadMarkdownStage(),
        MarkdownParseStage(),
        TreeBuildingStage(),
        DocumentCreationStage(),
        StyleDefinitionFixStage(),
        FormattingExecutionStage(),
    ]

    for stage in pipeline:
        ctx = stage.process(ctx)

    # 保存前 hook：编号/超链接等机制与预设统一在此执行（复用 before_document_save）
    from wordformat.hooks import hooks

    hooks.emit(
        "before_document_save",
        document=ctx.document,
        config=ctx.config_model,
        ctx=ctx,
        check=False,
    )

    # 重命名输出文件
    ensure_directory_exists(save_dir)
    filename = get_file_name(md_path)
    out_path = Path(save_dir) / f"{filename}--生成版.docx"
    ctx.document.save(str(out_path))
    logger.info(f"保存文件到 {out_path}")
    return str(out_path)
