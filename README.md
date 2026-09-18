# 郭少技能库 · guoshaoskill

收集可复用的AI技能，将具体任务的输入要求、处理流程、输出格式和检查方法整理成可安装的技能包。每个技能均提供功能介绍、适用场景和新手教程，方便了解后再使用。

## 技能目录

| 技能 | 一句话用途 | 教程与下载 |
| --- | --- | --- |
| 课件转五篇笔记 | 将PDF课件改写为五种角度的中文笔记，每篇600字以内 | [教程](courseware-five-notes/README.md) · [下载](https://raw.githubusercontent.com/guoyaotian/guoshaoskill/main/courseware-five-notes/courseware-five-notes.zip) |
| 班会资料包 | 为已有班会课件配套生成课堂逐字稿与教案，分别保存为Word | [教程](class-meeting-packager/README.md) · [下载](https://raw.githubusercontent.com/guoyaotian/guoshaoskill/main/class-meeting-packager/class-meeting-packager.zip) |

## 📚 课件转五篇笔记

**给它一份PDF课件，生成五篇不同角度的中文笔记，每篇不超过600字。** 适合教师、培训者、家长会组织者及教育内容创作者，用于介绍课件、梳理重点、准备社交媒体文案草稿。

| 你提供什么 | 它会做什么 | 你获得什么 |
| --- | --- | --- |
| PDF课件或可访问的文件路径；可选读者和语气偏好 | 阅读课件、核对页码、提炼真实内容，区分事实与建议，并检查字数 | 全貌版、教学亮点版、知识拆解版、应用场景版、使用建议版，共5篇Markdown笔记 |

主要功能：

- 按实际页码梳理内容，保留数据单位与语境。
- 五种角度分别写作，使用独立标题、emoji和分层排版。
- 不编造年龄、年级、时长或案例；新增使用方法明确标为建议。
- 提示原文矛盾、旧日期等需核对的信息。
- 附带脚本检查五篇的数量、顺序、非空正文及600字上限。

**[查看完整功能与新手教程](courseware-five-notes/README.md)** · **[直接下载技能包](https://raw.githubusercontent.com/guoyaotian/guoshaoskill/main/courseware-five-notes/courseware-five-notes.zip)** · [查看技能规则](courseware-five-notes/SKILL.md)

## 🏫 班会资料包

**给它一份班会课件，配套生成课堂逐字稿与教案两个 Word。** 适合中小学班主任、德育教师与教学资料整理者，用于主题班会、节气教育、常规教育等已有课件的配套备课。

| 你提供什么 | 它会做什么 | 你获得什么 |
| --- | --- | --- |
| PDF/PPT/PPTX课件；可选参考教案、年级、课时、字数与输出偏好 | 由宿主AI读课件并起草，附带脚本校验草稿并导出Word，再核对教学内容与版面 | 逐字稿与教案两个独立DOCX；也可明确只要一种 |

主要功能：

- 默认口播正文3500字以内，教师提示与正文分开，环节标题可标对应页码。
- 默认十章教案，含教学过程、整块板书、反思要点和教学评价。
- 检查字数、章节、页码范围、两份来源页覆盖以及导出文字完整性。
- 默认A4、真实标题层级、紧凑段距，可采样参考教案命名样式或自定义排版参数。
- 已有成品保留，同名时整组另存版本；清理限于本次任务的临时文件。

**[查看完整功能与新手教程](class-meeting-packager/README.md)** · **[直接下载技能包](https://raw.githubusercontent.com/guoyaotian/guoshaoskill/main/class-meeting-packager/class-meeting-packager.zip)** · [查看技能规则](class-meeting-packager/SKILL.md)

本包面向具备文档阅读和本地Python执行能力的Codex；导出器不自行读取课件或执行OCR。已通过38项回归测试，示例Word已渲染检查；复杂模板、其他宿主和任意真实课件尚未全面验证。详细依赖与限制见教程。

## 🚀 快速开始

1. 下载技能包，在所用软件中安装；详细步骤见技能教程。
2. 根据所选技能提供可读取的课件及可选偏好。
3. 课件转五篇笔记可发送：

```text
请使用 courseware-five-notes（课件转五篇笔记）技能，根据这份PDF生成5篇不同角度的笔记，每篇600字以内，核对页码并检查字数。
```

Codex也可使用 `$courseware-five-notes` 指定技能。本包已通过作者环境中的结构与脚本校验；WorkBuddy导入方式见教程，尚未实际试跑。PDF阅读与字数校验依赖宿主软件的文档工具和Python环境。

本技能不制作PPT、不自动发布内容，也不承诺流量效果。

班会资料包在Codex中可发送：

```text
使用 $class-meeting-packager 处理这份班会课件，输出3500字以内的课堂逐字稿与配套教案两个Word。保存到课件所在目录，核对内容、页码和版面；未完成的检查请明确说明。
```

## 技能介绍与更新规范

新增或更新技能时，应同步维护：**一句话用途、目标用户、输入材料、核心功能、输出结果、安装方法、调用示例、依赖及限制**。功能描述应以实际能力为准，源码、教程和ZIP安装包保持一致，避免读者下载到旧版。
