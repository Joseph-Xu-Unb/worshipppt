import { downloadUrl } from "../lib/app-config";

const projectInstructions = `Role: 你是一个专业的教会服事数据处理专家，擅长从杂乱的聊天记录中提取关键信息并转化为精准的 JSON 格式。

Task: 请阅读以下聊天记录，提取主日崇拜的经文和诗歌内容，并生成一个名为 worship_data.json 的文件。

Data Schema (JSON 结构要求):

call_to_worship: 包含 title（经文名字）和 lines（按圣经原文章节号分行，一节一行；行号不要加括号，如 "1 你们要..."）。

hymns: 一个数组，包含所有赞美诗和诗班献诗。每项含 title（需标注是否为诗班）和 lines（完整歌词）。三首赞美诗需按 4 句一页 PPT 的规则整理，可用空行作为分页分隔。诗歌中较长的句子应分割成 2 行，便于歌词以 48 号字体在 PPT 显示。

theme_scripture: 包含 title 和 lines（使用和合本，按圣经原文章节号分行，一节一行；行号不加括号）。

response_hymn: 包含 title 和 lines。

Formatting Rules (格式准则):

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
        <p className="eyebrow">Resources</p>
        <h2>Need Help Preparing The JSON?</h2>
        <p className="resource-copy">
          Download the sample data, then use ChatGPT or Muse AI to turn raw worship
          planning chats into the JSON this page expects.
        </p>
      </div>

      <div className="download-grid">
        <a className="download-card" href={downloadUrl("/downloads/template")}>
          <span className="download-label">Template</span>
          <strong>template.pptx</strong>
          <span>The same PowerPoint template from `data/` used by the server generator.</span>
        </a>
        <a className="download-card" href={downloadUrl("/downloads/sample-ppt")}>
          <span className="download-label">Sample Deck</span>
          <strong>2026-05-17.pptx</strong>
          <span>A real sample PowerPoint deck users can download for reference.</span>
        </a>
        <a className="download-card" href={downloadUrl("/downloads/sample-json")}>
          <span className="download-label">Example</span>
          <strong>Sample JSON</strong>
          <span>A ready example from `data/` showing the expected JSON structure.</span>
        </a>
      </div>

      <details className="guide-card">
        <summary>
          <span>ChatGPT Project Setup Guide</span>
          <span className="guide-summary-note">Open instructions</span>
        </summary>
        <h3>ChatGPT Project Setup</h3>
        <ol>
          <li>
            Open ChatGPT and create a new Project. Official guide:
            {" "}
            <a
              href="https://help.openai.com/en/articles/10169521-projects-in-chatgpt"
              target="_blank"
              rel="noreferrer"
            >
              Projects in ChatGPT
            </a>
          </li>
          <li>Upload the downloaded `template.pptx` and sample JSON into the project sources.</li>
          <li>
            In the project menu, open `Project settings`, then paste the instruction
            template below into project instructions.
          </li>
          <li>Paste your worship planning chat records into a project chat and ask ChatGPT to generate `worship_data.json`.</li>
          <li>Download the JSON result from ChatGPT and upload it here to preview and generate the final PPT.</li>
        </ol>
        <p className="guide-note">
          OpenAI says projects let you upload files and add project-specific
          instructions, and those instructions override your global custom
          instructions inside that project.
        </p>
        <textarea
          className="instructions-box"
          value={projectInstructions}
          readOnly
          aria-label="ChatGPT project instructions template"
        />
      </details>

      <details className="guide-card">
        <summary>
          <span>Muse AI Setup Guide</span>
          <span className="guide-summary-note">Open instructions</span>
        </summary>
        <h3>Use WorshipPPT with Muse AI</h3>
        <ol>
          <li>
            Open <a href="https://muse.ai/" target="_blank" rel="noreferrer">Muse AI</a>
            {" "}and start a conversation for your worship service.
          </li>
          <li>
            Ask Muse to connect to the WorshipPPT MCP server using Streamable HTTP
            at <code>https://worshipppt.onrender.com/mcp/</code>.
            WorshipPPT requires no sign-in, API key, or authorization header.
          </li>
          <li>
            Paste the connection prompt below. Provide your service date and worship
            planning notes, including the complete scripture and song text.
          </li>
          <li>
            Review the prepared content, then ask Muse to validate it and generate
            the PowerPoint. Ask Muse to save the returned presentation as a downloadable file.
          </li>
          <li>
            Review any generation warnings and open the downloaded PowerPoint to
            check the slides before the service.
          </li>
        </ol>
        <textarea
          className="instructions-box"
          value={`Connect to WorshipPPT using its MCP Streamable HTTP endpoint:
https://worshipppt.onrender.com/mcp/
Authentication: none. No API key or authorization header is needed.

Discover the available tools, then:
1. Call get_sample_worship_data to learn the required JSON structure.
2. Use my worship planning notes to prepare the data. Ask me for any missing service date, scripture, or lyrics. Do not invent missing content.
3. Show me the prepared content for review.
4. Call validate_worship_data and fix any validation errors.
5. When I ask you to generate the presentation, call generate_worship_powerpoint with data and service_date in YYYY-MM-DD format.
6. Decode pptx_base64 into the returned filename and offer the .pptx file for download. Show me any returned warnings.

Service date: [YYYY-MM-DD]
Worship planning notes: [paste notes here]`}
          readOnly
          aria-label="Muse AI MCP connection and presentation generation prompt"
        />
        <p className="guide-note">
          If Muse cannot connect to MCP or save the presentation, use the JSON workflow:
          paste the instruction template below into Muse, save its result as
          {" "}<code>worship_data.json</code>, then upload it on this app’s home page.
          Choose the service date, preview the slides, and generate the PowerPoint here.
        </p>
        <h3>Worship JSON Instructions</h3>
        <p className="guide-note">
          AI generated scripture and song lyrics can contain errors. Review the JSON
          against your approved Bible translation and source materials before creating
          the service deck.
        </p>
        <textarea
          className="instructions-box"
          value={projectInstructions}
          readOnly
          aria-label="Muse AI worship data instruction template"
        />
      </details>
    </section>
  );
}
