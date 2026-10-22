import { useState } from "react";

import { lyricsLibraryUrl } from "../lib/church-resources";

const musePrompt = `请把下面的教会群聊天记录整理成崇拜 PPT。
先列出聚会日期、宣召和主题经文、诗歌名称、完整歌词及顺序，供我核对。以群里最后确认的安排为准；有缺失或冲突先问我，不要编造。经文使用和合本，保留歌词重复段落和间奏。
教会歌词库：${lyricsLibraryUrl}
优先使用教会歌词库中的版本。若无法打开链接，请让我上传相应诗歌文件或粘贴完整歌词，不要凭诗歌名称猜写歌词。
我确认后，请生成可下载的 PowerPoint 文件。

供你连接制作服务：
MCP Streamable HTTP：https://worshipppt.onrender.com/mcp/
无需登录或 API 密钥。调用 get_sample_worship_data 查看格式，用 validate_worship_data 验证，再用 generate_worship_powerpoint 生成，service_date 使用 YYYY-MM-DD。将返回的 pptx_base64 解码为可下载的 .pptx，并说明 warnings。
如连接超时，先访问 https://worshipppt.onrender.com/api/health，等服务启动后重试。启动可能需要约一分钟。

以下是从教会微信群或 WhatsApp 群复制的聊天记录：
[请在这里粘贴聊天记录]`;

export function MuseGuide() {
  const [copyMessage, setCopyMessage] = useState("");

  async function copyPrompt() {
    try {
      await navigator.clipboard.writeText(musePrompt);
      setCopyMessage("已复制，请粘贴到 Muse AI。");
    } catch {
      setCopyMessage("请选中下方文字，手动复制。");
    }
  }

  return (
    <section className="resource-block">
      <div className="resource-header">
        <p className="eyebrow">Muse AI 使用指南</p>
        <h2>用教会群消息制作崇拜 PPT</h2>
        <p className="resource-copy">准备聊天记录，交给 Muse 整理，确认后下载 PPT。</p>
      </div>

      <div className="guide-card" style={{ padding: 18 }}>
        <h3>1. 复制教会群消息</h3>
        <p className="resource-copy">
          从微信或 WhatsApp 教会群复制本次崇拜的筹备消息，包括后续更正。
          请包含<strong>聚会日期、经文出处、诗歌名称和完整歌词</strong>，以及诗歌顺序。
          歌词若在图片或附件里，也请一起发给 Muse。
        </p>
        <p className="guide-note">
          <strong>缺少歌词？</strong>打开
          <a href={lyricsLibraryUrl} target="_blank" rel="noreferrer">教会歌词库（OneDrive）</a>，
          找到本次诗歌，把完整歌词复制给 Muse，或下载诗歌文件后上传给它。
        </p>
        <p className="guide-note">需要您手动提供这些消息，Muse 不会通过本应用自动读取教会群。</p>
      </div>

      <div className="guide-card" style={{ padding: 18 }}>
        <h3>2. 粘贴到 Muse AI</h3>
        <p className="resource-copy">
          打开 Muse AI：<a href="https://muse.ai/" target="_blank" rel="noreferrer">https://muse.ai/</a>。
          复制下面的文字，粘贴到新对话中，再把教会群消息加在最后并发送。
        </p>
        <button type="button" onClick={() => void copyPrompt()}>复制给 Muse 的文字</button>
        <p role="status" aria-live="polite">{copyMessage}</p>
        <details>
          <summary>查看要复制的文字</summary>
          <textarea className="instructions-box" value={musePrompt} readOnly aria-label="复制给 Muse AI 的文字" />
        </details>
      </div>

      <div className="guide-card" style={{ padding: 18 }}>
        <h3>3. 核对并下载 PPT</h3>
        <p className="resource-copy">
          检查 Muse 整理的日期、经文、歌词和顺序。资料齐全后，回复
          <strong>“确认，请生成 PPT”</strong>，再下载并打开文件检查排版。
        </p>
      </div>
      <p className="guide-note">
        <strong>连接不上？</strong>服务长时间未使用时可能需要约一分钟启动。
        请先<a href="https://worshipppt.onrender.com/api/health" target="_blank" rel="noreferrer">点这里唤醒服务</a>，
        等页面加载完成后，回到 Muse 说“请重试”。WorshipPPT 不需要登录或密钥。
      </p>
    </section>
  );
}
