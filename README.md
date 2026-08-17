# docx-spec-html Skill

这是 `docx-spec-html` 的完整发布仓库。可安装的 Skill 位于
`docx-spec-html/` 子目录；**不要只复制 `SKILL.md`**，否则会缺少
`scripts/`、`assets/` 和 `references/`。

## 交给 Agent 的安装指令

克隆完整仓库后，在仓库根目录运行：

### macOS

```bash
bash install.sh
```

### Windows PowerShell

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\install.ps1
```

安装器会：

1. 把完整 `docx-spec-html/` 复制到 JoyCode 的 Skill 目录；
2. 使用 `uv` 安装隔离的 Python 3.12；
3. 安装固定版本的 `python-docx`；
4. 从 `iOfficeAI/OfficeCLI` 官方发行渠道安装 OfficeCLI；
5. 检查全部必要脚本、资源和运行依赖；
6. 仅在显示 `READY` 后报告安装成功。

默认安装位置：

- macOS：`~/.joycode/skills/docx-spec-html`
- Windows：`%USERPROFILE%\.joycode\skills\docx-spec-html`

安装后重启 JoyCode，使其重新加载 Skill。

## 自检

macOS：

```bash
~/.joycode/skills/docx-spec-html/.venv/bin/python \
  ~/.joycode/skills/docx-spec-html/scripts/preflight.py
```

Windows：

```powershell
& "$env:USERPROFILE\.joycode\skills\docx-spec-html\.venv\Scripts\python.exe" `
  "$env:USERPROFILE\.joycode\skills\docx-spec-html\scripts\preflight.py"
```
