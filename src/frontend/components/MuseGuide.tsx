import { ChatHistoryInput } from "./ChatHistoryInput";

import { lyricsLibraryUrl } from "../lib/church-resources";

const musePrompt = `请把下面的教会群聊天记录整理成崇拜 PPT。
从聊天记录提取聚会日期、宣召和主题经文、诗歌名称、完整歌词及顺序，以群里最后确认的安排为准。经文使用和合本，保留歌词重复段落和间奏。资料完整时直接制作，不需要我检查 JSON 或再次确认；只有必要资料缺失或消息冲突时才询问，不要编造。
教会歌词库：${lyricsLibraryUrl}
优先使用教会歌词库中的版本。读取歌词库时，请按以下步骤操作：
1. 优先使用你可用且获准使用的浏览器工具打开 OneDrive 链接，等待页面中的文件列表加载，再查找并打开本次诗歌，读取完整歌词；不要只依赖直接抓取页面文本，因为 OneDrive 页面需要 JavaScript 渲染。
2. 如果轻量文本抓取失败或提示域名受限，不要据此认定链接失效、文件不存在或需要登录；如果有获准使用的浏览器工具，请用浏览器打开原链接。遵守工具的访问限制，不要绕过访问控制。
3. 此链接为公开分享，先尝试匿名访问，不要预先要求我登录。如果页面提示渲染失败或暂时出错，请在新标签页重新打开原链接，等待加载后重试一次，避免反复刷新。
4. 找到诗歌后，核对标题及版本并读取全部歌词，保留重复段落和段落标记，不要只根据文件名、搜索摘要或部分预览补写。读取时记录“间奏”“前奏”“尾声”“结束”等标记对应的歌词段落和原幻灯片页码，不要只记录“有间奏”而丢失位置。
5. 只有在没有可用的浏览器工具、浏览器重试后仍无法读取、页面确实要求登录或没有找到所需诗歌时，才说明具体原因，并请我上传相应诗歌文件或粘贴完整歌词。不要凭诗歌名称猜写歌词。
请自动整理资料、验证并调用制作服务，把最终可下载的 PowerPoint 文件直接交给我。JSON 仅供工具内部使用，不要要求我查看、下载或上传 JSON。
歌词标记排版规则：“间奏”“前奏”“尾声”“结束”等诗歌段落标记统一放在对应歌词页的末尾、该页最后一句歌词之后，不得移到下一页或下一段的开头，也不得单独占一页。位置不明确时，归到前一段歌词所在页的末尾；若无法确定对应段落，请询问，不要猜测。分页时须让标记与对应歌词保持在同一页。此规则针对诗歌中的标记，不影响崇拜流程的结束页。
交付前请检查实际生成的 PPTX：如有文件处理工具，解压并按幻灯片顺序读取每页文字，核对歌词完整性、重复段落、分页以及上述标记是否位于对应歌词页末尾；发现错位时修正资料并重新生成、复查。不要仅凭制作服务返回“生成成功”就认定排版正确。如工具无法检查文件，请明确说明未完成文件检查，不要声称已核验。

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
          也可以让 Muse 按下方指令先用浏览器查找本次诗歌。若仍无法读取，再把完整歌词复制给 Muse，或下载诗歌文件后上传给它。
        </p>
        <p className="guide-note">
          歌词库为公开分享，可先尝试匿名访问。若 Muse 的文本抓取失败，请让它使用可用的浏览器工具打开链接并等待加载；
          若 OneDrive 提示渲染失败，可在新标签页重新打开一次。文本抓取失败不一定表示需要登录。
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
