# tampermonkey-scripts

自用的油猴（Tampermonkey）脚本合集，主要用来导出会议、聊天和文章内容。

## 安装

1. 浏览器里装好 [Tampermonkey](https://www.tampermonkey.net/)
2. 打开下表里的脚本文件，点 **Raw**，Tampermonkey 会弹出安装页面

## 脚本一览

### 自己写的

| 脚本 | 适用网站 | 功能 |
|---|---|---|
| [Teams聊天导出](Teams聊天导出.user.js) | teams.cloud.microsoft | 按时间范围导出对话，保存为 HTML（图片内嵌）或 Markdown ZIP（含图片、附件）；支持勾选多个对话批量导出 |
| [SharePoint会议转写获取](SharePoint会议转写获取.user.js) | SharePoint / Stream 会议录像 | 获取完整的会议转写，可选 TXT / TSV / JSON 保存 |
| [微信公众号转Markdown](微信公众号转Markdown.user.js) | mp.weixin.qq.com | 把公众号文章转成 Markdown，连同图片打包成 ZIP |

Teams 聊天导出基于 Greasy Fork 上的 [Teams Chat Exporter](https://greasyfork.org/scripts/537462)（MIT）大幅改写。

### 收录的第三方脚本

以下脚本原样收录，版权和许可证归原作者所有，最新版本请到原地址获取。

| 脚本 | 作者 | 许可证 | 来源 |
|---|---|---|---|
| [ChatGPT对话导出](ChatGPT对话导出.user.js) | pionxzh | MIT | [Greasy Fork](https://greasyfork.org/scripts/456055) |
| [Gemini对话备份](Gemini对话备份.user.js) | Sxuan-Coder / AXIRYYYY | Apache-2.0 | [GitHub](https://github.com/AXIRYYYY/gemini-chat-export-enhanced) |
| [知乎备份剪藏](知乎备份剪藏.user.js) | qtqz | MIT | [GitHub](https://github.com/qtqz/zhihu-backup-collect) |
| [腾讯会议字幕导出](腾讯会议字幕导出.user.js) | — | MIT | [Greasy Fork](https://greasyfork.org/scripts/547992) |
| [腾讯会议转写纪要导出](腾讯会议转写纪要导出.user.js) | 东哥说AI | MIT | [Greasy Fork](https://greasyfork.org/scripts/540670) |

## 许可证

自己写的脚本以 [MIT](LICENSE) 许可证发布。第三方脚本遵循各自的许可证（见上表和脚本头部）。
