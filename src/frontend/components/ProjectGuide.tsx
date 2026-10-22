import { downloadUrl } from "../lib/app-config";

const projectInstructions = `角色： 你是一个专业的教会服事数据处理专家，擅长从杂乱的聊天记录中提取关键信息并转化为精准的 JSON 格式。

任务： 请阅读以下聊天记录，提取主日崇拜的经文和诗歌内容，并生成一个名为 worship_data.json 的文件。

先整理并请用户确认：聚会日期（YYYY-MM-DD）、宣召经文出处、主题经文出处、诗歌名称、完整歌词及崇拜顺序。同一事项有更正时，以最后明确确认的信息为准；存在冲突或缺失时先询问用户，不要编造经文或歌词。聚会日期需在本应用中单独选择，不必添加到 JSON 中。

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

export function ProjectGuide({ provider = "chatgpt" }: { provider?: "muse" | "chatgpt" }) {
  const isMuse = provider === "muse";
  return (
    <section className="resource-block">
      <div className="resource-header">
        <p className="eyebrow">使用指南</p>
        <h2>{isMuse ? "Muse AI：从教会群聊天记录制作崇拜 PPT" : "ChatGPT：从教会群聊天记录制作崇拜 PPT"}</h2>
        <p className="resource-copy">
          先从教会微信群或 WhatsApp 群复制与本次崇拜 PPT 有关的聊天记录，
          再使用 {isMuse ? "Muse AI" : "ChatGPT"} 整理内容并制作简报。
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

      {isMuse && <details className="guide-card" open>
        <summary>
          <span>Muse AI 使用指南</span>
          <span className="guide-summary-note">展开说明</span>
        </summary>
        <h3>使用 Muse AI 制作崇拜 PPT</h3>
        <ol>
          <li>打开 <a href="https://muse.ai/" target="_blank" rel="noreferrer">Muse AI</a>，开始一次崇拜筹备对话。</li>
          <li>复制下方提示词，并粘贴从教会微信群或 WhatsApp 群收集的聊天记录。重点提供聚会日期、经文出处、诗歌名称和完整歌词。</li>
          <li>请 Muse 通过 Streamable HTTP 连接 <code>https://worshipppt.onrender.com/mcp/</code>。WorshipPPT 服务不需要登录、API 密钥或授权请求头。</li>
          <li>先核对 Muse 整理的日期、经文、歌词和顺序。确认后，再请它验证数据并生成可下载的 PowerPoint 文件。</li>
          <li>查看返回的提示信息，并打开下载的 PPT，检查文字和排版后再用于聚会。</li>
        </ol>
        <textarea
          className="instructions-box"
          value={`请根据我从教会微信群或 WhatsApp 群复制的崇拜 PPT 筹备聊天记录，协助制作本次主日崇拜简报。

通过 MCP Streamable HTTP 连接 WorshipPPT：
https://worshipppt.onrender.com/mcp/
本服务无需登录、API 密钥或授权请求头。

请依次完成：
1. 发现可用工具，调用 get_sample_worship_data 了解 JSON 数据结构。
2. 从聊天记录提取聚会日期、宣召经文出处及正文、主题经文出处及正文、赞美诗/诗班献诗/回应诗的名称、完整歌词和顺序。经文使用和合本；保留歌词重复段落和间奏。
3. 同一事项有更正时，以最后明确确认的信息为准。日期不清楚、内容缺失或消息冲突时，先问我，不要编造经文或歌词。
4. 列出整理结果供我核对，尤其是聚会日期、经文章节范围、歌词完整性和诗歌顺序。
5. 调用 validate_worship_data 验证数据，并修正格式错误。
6. 等我确认生成后，调用 generate_worship_powerpoint，将 data 和 YYYY-MM-DD 格式的 service_date 传入。
7. 将返回的 pptx_base64 解码，按返回的 filename 保存为可下载的 .pptx 文件，并向我说明 warnings 中的提示。

聚会日期：[YYYY-MM-DD；如未确定，请从聊天记录整理后向我确认]
教会群聊天记录（包括补充、更正及完整歌词）：
[在这里粘贴微信或 WhatsApp 群里的相关消息]`}
          readOnly
          aria-label="Muse AI 连接服务及制作崇拜 PPT 的中文提示词"
        />
        <p className="guide-note">
          如果 Muse 无法连接 MCP 或保存 PPT，请使用下方的 JSON 指令整理聊天记录，
          将结果保存为 <code>worship_data.json</code>，再回到本应用首页上传文件、
          选择聚会日期、预览幻灯片并生成 PPT。
        </p>
        <h3>整理崇拜 JSON 的指令</h3>
        <p className="guide-note">请根据教会确认的经文版本和歌词原文核对 AI 整理的内容。</p>
        <textarea className="instructions-box" value={projectInstructions} readOnly aria-label="Muse AI 崇拜数据中文指令模板" />
      </details>}

      {!isMuse && <details className="guide-card" open>
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
      </details>}
    </section>
  );
}
