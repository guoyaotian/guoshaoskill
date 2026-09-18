# 郭少技能库

可复用的技能、安装包和新手使用教程。

## 课件转五篇笔记

根据PDF课件生成五篇不同角度的中文笔记，每篇不超过600字：课程全貌、教学亮点、知识拆解、应用场景、使用建议。包括页码核对、原文事实与建议区分，以及字数校验脚本。

- [查看技能说明](courseware-five-notes/SKILL.md)
- [阅读新手教程（Codex / WorkBuddy）](courseware-five-notes/README.md)
- [下载技能安装包](courseware-five-notes/courseware-five-notes.zip)

安装包页面如未直接下载，点击 GitHub 的下载原始文件按钮即可。

### Codex 调用示例

```text
用 $courseware-five-notes，严格根据这份PDF生成5篇不同角度的笔记，每篇600字以内，核对页码并检查字数。
```

### WorkBuddy 调用示例

```text
请使用 courseware-five-notes（课件转五篇笔记）技能，根据我提供的PDF生成5篇不同角度的笔记，每篇600字以内，并检查字数。
```

WorkBuddy支持本地技能包导入；本包尚未在WorkBuddy中实际试跑。PDF读取与自动字数校验需要所用软件提供相应的文档处理和Python运行环境。
