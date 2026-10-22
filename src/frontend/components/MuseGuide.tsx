import { ChatHistoryInput } from "./ChatHistoryInput";

import { lyricsLibraryUrl } from "../lib/church-resources";

const musePrompt = `请把下面的教会群聊天记录整理成崇拜 PPT。
从聊天记录提取聚会日期、宣召和主题经文、诗歌名称、完整歌词及顺序，以群里最后确认的安排为准。经文使用和合本，保留歌词重复段落和间奏。资料完整时直接制作，不需要我检查 JSON 或再次确认；只有必要资料缺失或消息冲突时才询问，不要编造。
教会歌词库：${lyricsLibraryUrl}
优先使用教会歌词库中的版本。若无法打开链接，请让我上传相应诗歌文件或粘贴完整歌词，不要凭诗歌名称猜写歌词。
请自动整理资料、验证并调用制作服务，把最终可下载的 PowerPoint 文件直接交给我。JSON 仅供工具内部使用，不要要求我查看、下载或上传 JSON。

供你连接制作服务：
MCP Streamable HTTP：https://worshipppt.onrender.com/mcp/
无需登录或 API 密钥。调用 get_sample_worship_data 查看格式，用 validate_worship_data 验证，再用 generate_worship_powerpoint 生成，service_date 使用 YYYY-MM-DD。将返回的 pptx_base64 解码为可下载的 .pptx，并说明 warnings。
如连接超时，先访问 https://worshipppt.onrender.com/api/health，等服务启动后重试。启动可能需要约一分钟。`;

export function MuseGuide() {
  return (
    <section className="resource-block">
      <div className="resource-header">
        <p className="eyebrow">Muse AI 使用指南</p>
        <h2>用教会群消息制作崇拜 PPT</h2>
        <p className="resource-copy">粘贴聊天记录，发给 Muse，直接获取最终 PPT。</p>
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
          在下方粘贴教会群消息，点击“复制指令和聊天记录”，再粘贴到 Muse 的新对话中发送一次。
          Muse 会通过 WorshipPPT 制作服务生成最终文件，您无需查看 JSON。
        </p>
        <ChatHistoryInput instructions={musePrompt} assistant="Muse AI" />
      </div>

      <div className="guide-card" style={{ padding: 18 }}>
        <h3>3. 下载最终 PPT</h3>
        <p className="resource-copy">
          Muse 会自动整理资料并生成 PPT，完成后直接下载文件。
          如果必要资料缺失，它会请您补充；资料齐全时，无需再次确认即可制作。
        </p>
      </div>
      <p className="guide-note">
        请求需要在 Muse 中发送。本页按钮帮您复制内容，不会自动操作 Muse。
        WorshipPPT 不需要另外配置 OpenAI API 密钥。
      </p>
      <p className="guide-note">
        <strong>连接不上？</strong>服务长时间未使用时可能需要约一分钟启动。
        请先<a href="https://worshipppt.onrender.com/api/health" target="_blank" rel="noreferrer">点这里唤醒服务</a>，
        等页面加载完成后，回到 Muse 说“请重试”。WorshipPPT 不需要登录或密钥。
      </p>
    </section>
  );
}
