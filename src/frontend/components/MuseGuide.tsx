import { ChatHistoryInput } from "./ChatHistoryInput";

import { lyricsLibraryUrl, worshipArchiveUrl } from "../lib/church-resources";
import musePromptTemplate from "../../shared/muse_prompt.txt?raw";

const musePrompt = musePromptTemplate
  .split("{{LYRICS_LIBRARY_URL}}")
  .join(lyricsLibraryUrl)
  .split("{{WORSHIP_ARCHIVE_URL}}")
  .join(worshipArchiveUrl);

export function MuseGuide() {
  return (
    <section className="resource-block">
      <div className="resource-header">
        <p className="eyebrow">Muse AI 使用指南</p>
        <h2>用教会群消息制作崇拜 PPT</h2>
        <p className="resource-copy">粘贴聊天记录，Muse 默认复制上周事项并写回 OneDrive，完成后返回可编辑分享链接。</p>
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
          <strong>事项汇报与写回：</strong>默认不加询问，Muse 会自动从
          <a href={worshipArchiveUrl} target="_blank" rel="noreferrer">往期崇拜 PPT 归档库（OneDrive）</a>
          读取上周文件并复制事项页（文字、图片、表格、分组形状和图表），然后把新 PPT 按
          &lt;本次日期&gt;.pptx 写回归档，返回可编辑分享链接和下载链接。
          只有找不到上周事项汇报时，Muse 才会停下来问：是继续生成（事项页留空），还是由您提供上周文件/页码。
        </p>
        <p className="guide-note">
          <strong>分享与 review：</strong>写回成功后，Muse 会把
          <code>uploaded_edit_url</code> 完整显示给您。请把该链接转发到教会群聊，请牧师和主席打开链接
          review，并直接在这个共享链接上继续修改。<code>uploaded_share_url</code> 是无需登录的下载链接，也请完整保留。
          如果可编辑链接无法匿名使用，请让牧师/主席登录教会 OneDrive 后在归档文件夹中编辑。
        </p>
        <p className="guide-note">
          <strong>周日使用：</strong>只要惠普（HP）和 DELL 电脑同步了该 OneDrive 文件夹，周日直接打开本地
          OneDrive 的 PPT 文件夹，找到 &lt;本次日期&gt;.pptx 打开使用即可，不必再找分享链接。
        </p>
        <p className="guide-note">
          <strong>新增诗歌：</strong>需要新增诗歌时，可以让 Muse 把新歌保存到
          <a href={lyricsLibraryUrl} target="_blank" rel="noreferrer">教会 OneDrive 曲库</a>。
          保存前先确认歌名、歌词和版本，不要覆盖或改动已有诗歌文件；WorshipPPT 只负责生成 PPT，不会自动写曲库。
        </p>
        <p className="guide-note">
          歌词库及归档库为公开分享，可先尝试匿名访问。若 Muse 的文本抓取失败，请让它使用可用的浏览器工具打开链接并等待加载；
          若 OneDrive 提示渲染失败，可在新标签页重新打开一次。文本抓取失败不一定表示需要登录。
        </p>
        <p className="guide-note">需要您手动提供这些消息，Muse 不会通过本应用自动读取教会群。</p>
      </div>

      <div className="guide-card" style={{ padding: 18 }}>
        <h3>2. 粘贴到 Muse AI</h3>
        <p className="resource-copy">
          打开 Muse AI：<a href="https://muse.ai/" target="_blank" rel="noreferrer">https://muse.ai/</a>。
          在下方粘贴教会群消息，点击“复制指令和聊天记录”，再粘贴到 Muse 的新对话中发送一次。
          Muse 会默认复制上周事项、生成 PPT 并写回 OneDrive；找不到上周事项页时才会先问您。您无需查看 JSON。
        </p>
        <ChatHistoryInput instructions={musePrompt} assistant="Muse AI" />
        <p className="guide-note">
          <strong>提示词同步：</strong>Muse 每次调用制作服务前会先检查提示词版本；如果 Muse
          保存的提示词过期，服务会拒绝生成并返回最新提示词，Muse 必须先更新自己的指令再重试。
        </p>
      </div>

      <div className="guide-card" style={{ padding: 18 }}>
        <h3>3. 分享、review 与本地使用</h3>
        <p className="resource-copy">
          生成并写回成功后，Muse 会把完整的可编辑分享链接（<code>uploaded_edit_url</code>）和无需登录的下载链接（
          <code>uploaded_share_url</code>）发给您。请把可编辑链接转发到教会群聊，请牧师和主席打开链接 review，
          并直接在这个共享链接上继续修改。周日打开惠普（HP）和 DELL 电脑上的本地 OneDrive 文件夹，
          找到 &lt;本次日期&gt;.pptx 打开使用即可。如果写回失败或暂缓，Muse 会说明原因，必要时仍把可下载的 PPT 交给您。
          资料齐全时不需要检查 JSON。
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
