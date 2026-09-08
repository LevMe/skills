# LevMe Skills

这是一个用于共享个人 Codex skills 的仓库。每个 skill 使用独立目录保存，目录名与 skill 名称一致；skill 自身的说明、规则、资源和脚本均放在对应目录内。

## 当前 skill

| 目录 | 名称 | 用途 |
| --- | --- | --- |
| [chinese-docx-format](./chinese-docx-format/) | 标准中文文档格式 | 控制中文 DOCX 的字体、版式、标题、图表、公式、引用、参考文献、附录和页码等非内容格式 |

## 安装

将需要使用的 skill 目录复制到 Codex 用户技能目录：

Windows：

    C:\Users\你的用户名\.codex\skills\<skill-name>

macOS 或 Linux：

    ~/.codex/skills/<skill-name>

安装后重启 Codex 或新建任务，使 skill 被重新发现。

## 使用

优先使用 skill 的显式调用名，或按照该 skill 目录中的 README.md 使用自然语言触发词。每个 skill 的具体执行规则以对应目录中的 SKILL.md 为准。

## 目录约定

每个 skill 至少包含一个 SKILL.md。如果需要自动调用元数据、参考资料、模板或脚本，应继续放在该 skill 的独立目录中，不要混入仓库根目录。

## 许可证

本仓库采用 Apache License 2.0，详见 [LICENSE](./LICENSE)。
