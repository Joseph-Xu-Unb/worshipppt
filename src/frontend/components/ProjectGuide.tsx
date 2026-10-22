import { downloadUrl } from "../lib/app-config";
import { lyricsLibraryUrl } from "../lib/church-resources";

const projectInstructions = `角色： 你是一个专业的教会服事数据处理专家，擅长从杂乱的聊天记录中提取关键信息并转化为精准的 JSON 格式。

任务： 请阅读以下聊天记录，提取主日崇拜的经文和诗歌内容，并生成一个名为 worship_data.json 的文件。

先整理并请用户确认：聚会日期（YYYY-MM-DD）、宣召经文出处、主题经文出处、诗歌名称、完整歌词及崇拜顺序。同一事项有更正时，以最后明确确认的信息为准；存在冲突或缺失时先询问用户，不要编造经文或歌词。聚会日期需在本应用中单独选择，不必添加到 JSON 中。

教会歌词库：${lyricsLibraryUrl}
优先使用教会歌词库中的版本。若无法打开链接，请让用户上传相应诗歌文件或粘贴完整歌词，不要凭诗歌名称猜写歌词。

JSON 结构要求：

call_to_worship: 包含 title（经文名字）和 lines（按圣经原文章节号分行，一节一行；行号不要加括号，如 "1 你们要..."）。

hymns: 一个数组，包含所有赞美诗和诗班献诗。每项含 title（需标注是否为诗班）和 lines（完整歌词）。三首赞美诗需按 4 句一页 PPT 的规则整理，可用空行作为分页分隔。诗歌中较长的句子应分割成 2 行，便于歌词以 48 号字体在 PPT 显示。

theme_scripture: 包含 title 和 lines（使用和合本，按圣经原文章节号分行，一节一行；行号不加括号）。

response_hymn: 包含 title 和 lines。

格式准则：

经文版本：必须使用“和合本”。

经文分行规则：经文必须根据圣经原文章节号分行，即一节一行；每节开头的数字后面直接跟内容，禁止使用 (1) 或 [1]。

歌词完整性：必须包含聊天记录中提到的所有歌词，包括重复段落和标注“间奏”的地方。

诗歌分页规则：三首赞美诗在生成 PPT 时必须按 4 句歌词一页进行分页；回应诗按实际歌词结构整理。

逻辑顺序：按照聊天中确认的顺序排列赞美诗（通常是赞美诗 1 -> 诗班献诗 1 -> 诗班献诗 2）。

纯文本输出：只输出符合 JSON 语法的 .json 文件。`;

export function ProjectGuide() {
  return (
    <section className="resource-block">
      <div className="resource-header">
        <p className="eyebrow">使用指南</p>
        <h2>ChatGPT：从教会群聊天记录制作崇拜 PPT</h2>
        <p className="resource-copy">
          先从教会微信群或 WhatsApp 群复制与本次崇拜 PPT 有关的聊天记录，
          再使用 ChatGPT 整理内容并制作简报。
        </p>
      </div>

      <section className="guide-card" aria-labelledby="chat-preparation-title">
        <div style={{ padding: 18 }}>
          <h3 id="chat-preparation-title">第一步：收集教会群里的崇拜资料</h3>
          <ol>
            <li>找到本次主日崇拜的筹备消息，复制相关聊天记录，包括后续补充和更正。</li>
            <li><strong>聚会日期：</strong>确认消息对应哪一次主日，整理为 YYYY-MM-DD 格式。</li>
            <li><strong>经文出处：</strong>收集宣召经文、主题经文的书卷名称和章节范围，以及群里提供的经文正文。</li>
            <li><strong>诗歌资料：</strong>收集赞美诗、诗班献诗及回应诗的名称、完整歌词、重复段落、间奏和演唱顺序。</li>
            <li>歌词或经文若在图片、附件或不同消息中，请一并提供给 AI，并核对识别出的文字。资料不全或有冲突时，先向同工确认。</li>
          </ol>
          <p className="guide-note">
            <strong>缺少歌词？</strong>打开
            <a href={lyricsLibraryUrl} target="_blank" rel="noreferrer">教会歌词库（OneDrive）</a>，
            找到本次诗歌，把完整歌词复制给 ChatGPT，或下载诗歌文件后上传给它。
            请使用教会确认的歌词版本。
          </p>
          <p className="guide-note">
            请手动复制并提供聊天记录；本应用不会自动读取微信或 WhatsApp 群。
            只提供诗歌名称时，还需要补充完整歌词。把相关消息一起交给 AI，有助于它识别最终确认的安排。
          </p>
        </div>
      </section>

      <div className="download-grid">
        <a className="download-card" href={downloadUrl("/downloads/template")}>
          <span className="download-label">PPT 模板</span>
          <strong>template.pptx</strong>
          <span>本应用制作崇拜简报时使用的模板。</span>
        </a>
        <a className="download-card" href={downloadUrl("/downloads/sample-ppt")}>
          <span className="download-label">简报示例</span>
          <strong>2026-05-17.pptx</strong>
          <span>下载已有的崇拜简报，参考版式和内容安排。</span>
        </a>
        <a className="download-card" href={downloadUrl("/downloads/sample-json")}>
          <span className="download-label">数据示例</span>
          <strong>示例 JSON</strong>
          <span>查看本应用需要的经文和歌词数据格式。</span>
        </a>
      </div>

      <details className="guide-card" open>
        <summary>
          <span>ChatGPT 项目设置指南</span>
          <span className="guide-summary-note">展开说明</span>
        </summary>
        <h3>使用 ChatGPT 整理教会群聊天记录</h3>
        <ol>
          <li>打开 ChatGPT，创建新项目。可参考<a href="https://help.openai.com/en/articles/10169521-projects-in-chatgpt" target="_blank" rel="noreferrer">官方项目指南</a>。</li>
          <li>下载上方的示例 JSON，并上传至项目作为格式参考。</li>
          <li>将下方中文指令粘贴到项目指令中。</li>
          <li>在项目对话中粘贴教会微信群或 WhatsApp 群的崇拜筹备记录，包括聚会日期、经文出处、完整歌词和后续更正。</li>
          <li>核对 AI 整理的资料，请它生成 <code>worship_data.json</code>，下载后在本应用首页上传，并单独选择已确认的聚会日期。</li>
          <li>预览幻灯片，核对内容和顺序后生成 PPT。</li>
        </ol>
        <textarea className="instructions-box" value={projectInstructions} readOnly aria-label="ChatGPT 项目中文指令模板" />
      </details>
    </section>
  );
}
