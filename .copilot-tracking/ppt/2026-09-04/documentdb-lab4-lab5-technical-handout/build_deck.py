from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

OUT_DIR = Path(__file__).parent / "slide-deck"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_PATH = OUT_DIR / "documentdb-lab4-lab5-technical-handout.pptx"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

COLORS = {
    "ink": RGBColor(31, 41, 55),
    "muted": RGBColor(75, 85, 99),
    "bg": RGBColor(247, 249, 252),
    "navy": RGBColor(20, 49, 86),
    "teal": RGBColor(0, 119, 130),
    "green": RGBColor(36, 125, 84),
    "amber": RGBColor(180, 83, 9),
    "red": RGBColor(185, 28, 28),
    "line": RGBColor(148, 163, 184),
    "white": RGBColor(255, 255, 255),
    "soft_teal": RGBColor(218, 244, 246),
    "soft_green": RGBColor(226, 246, 235),
    "soft_amber": RGBColor(255, 242, 214),
    "soft_blue": RGBColor(225, 239, 255),
    "soft_red": RGBColor(254, 226, 226),
}


def set_run(run, size=20, bold=False, color="ink", font="Aptos"):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = COLORS[color]


def add_text(slide, text, x, y, w, h, size=20, bold=False, color="ink", align=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    p = tf.paragraphs[0]
    if align:
        p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)
    return box


def add_title(slide, title, subtitle=None, section=None):
    if section:
        add_text(slide, section.upper(), 0.55, 0.25, 3.2, 0.3, size=10, bold=True, color="teal")
    add_text(slide, title, 0.55, 0.56, 8.6, 0.65, size=27, bold=True, color="navy")
    if subtitle:
        add_text(slide, subtitle, 0.57, 1.12, 9.8, 0.38, size=13, color="muted")
    line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.55), Inches(1.55), Inches(12.2), Inches(0.03))
    line.fill.solid()
    line.fill.fore_color.rgb = COLORS["line"]
    line.line.fill.background()


def add_box(slide, text, x, y, w, h, fill="soft_blue", line="line", color="ink", size=14, bold=False):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLORS[fill]
    shape.line.color.rgb = COLORS[line]
    shape.line.width = Pt(1)
    tf = shape.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.12)
    tf.margin_right = Inches(0.12)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)
    return shape


def add_bullets(slide, items, x, y, w, h, size=15, color="ink", gap=0.1):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.space_after = Pt(gap * 18)
        p.font.name = "Aptos"
        p.font.size = Pt(size)
        p.font.color.rgb = COLORS[color]
    return box


def add_arrow(slide, x1, y1, x2, y2, color="line"):
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    line.line.color.rgb = COLORS[color]
    line.line.width = Pt(1.6)
    return line


def add_notes(slide, notes):
    tf = slide.notes_slide.notes_text_frame
    tf.clear()
    tf.text = notes.strip()


def new_slide(title, subtitle=None, section=None, notes=None):
    slide = prs.slides.add_slide(BLANK)
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = COLORS["bg"]
    add_title(slide, title, subtitle, section)
    if notes:
        add_notes(slide, notes)
    return slide


def add_timeline(slide, labels, y=6.68):
    x = 0.7
    width = 11.9 / len(labels)
    for idx, label in enumerate(labels):
        fill = "soft_teal" if idx % 2 == 0 else "soft_green"
        add_box(slide, label, x + idx * width, y, width - 0.08, 0.42, fill=fill, size=9, bold=True)


# Slide 1
slide = prs.slides.add_slide(BLANK)
slide.background.fill.solid()
slide.background.fill.fore_color.rgb = RGBColor(235, 245, 247)
add_text(slide, "Azure DocumentDB", 0.65, 0.55, 6.2, 0.55, size=18, bold=True, color="teal")
add_text(slide, "Technical walkthrough speaker handout", 0.65, 1.28, 8.8, 0.75, size=30, bold=True, color="navy")
add_text(slide, "Lab 4 Search + Lab 5 RAG Pipeline | 30 minutes each | BM25 and hybrid treated as architecture walkthroughs when preview features are unavailable", 0.68, 2.18, 10.9, 0.55, size=15, color="muted")
add_box(slide, "Trainer stance\nExplain the system as an engineer would debug it", 0.75, 3.35, 3.5, 1.1, fill="soft_teal", bold=True)
add_box(slide, "Lab 4\nVector search works; BM25/hybrid may be preview-gated", 4.95, 3.35, 3.5, 1.1, fill="soft_amber", bold=True)
add_box(slide, "Lab 5\nRAG is the composition pattern, not a single query feature", 9.15, 3.35, 3.5, 1.1, fill="soft_green", bold=True)
add_notes(slide, "Open by framing this as a technical walkthrough rather than a click-through lab. The audience is highly technical, so treat every step as a design decision: identity boundary, data shape, embedding boundary, index boundary, retrieval boundary, and generation boundary. Be explicit that BM25 and hybrid search may be gated or unavailable in this environment. The goal is still strong: attendees should leave knowing where those features fit, how the code would use them, and how to validate the fallback path.")

# Slide 2
slide = new_slide("60-minute delivery map", "Use the clock aggressively; the labs are demos plus architecture, not full typing exercises.", "Agenda", "This is your pacing slide. Explain that Lab 4 gets 30 minutes and Lab 5 gets 30 minutes. For Lab 4, spend most time on setup, embeddings, vector indexing, and the preview-gated search story. For Lab 5, move quickly from chunking to retrieval to prompt assembly. If something fails, do not spend the room's time debugging package installs. Move to the conceptual path and show the relevant notebook cells.")
add_timeline(slide, ["0-5\nFrame", "5-15\nLab 4 setup + data", "15-25\nVector + preview search", "25-30\nLab 4 recap", "30-40\nLab 5 pipeline", "40-52\nRetrieval + prompt", "52-60\nTradeoffs + Q&A"], y=2.0)
add_bullets(slide, [
    "Do not promise BM25 or hybrid execution when full-text preview is unavailable.",
    "Keep the audience oriented around contracts: env vars, token scopes, collection names, index names, query shapes.",
    "Use failed preview queries as an opportunity to explain capability detection and graceful fallback.",
    "End each lab with the production design question: what would you monitor, tune, or change?",
], 1.0, 3.05, 11.2, 2.1, size=17)

# Slide 3
slide = new_slide("Shared runtime contract", "Lab 4 and Lab 5 use the same identity and endpoint contract.", "Architecture", "Use this slide before opening either notebook. The key message is that students should not edit cluster names in notebooks. The VM setup and Set-LabEnvironment.ps1 populate three environment variables. The notebooks consume them. Azure CLI supplies Entra tokens for both DocumentDB and Azure OpenAI. This is a cleaner model than copying connection strings or keys into notebooks.")
add_box(slide, "Trainer / student\nPowerShell + az login", 0.8, 2.1, 2.15, 0.9, fill="soft_blue", bold=True)
add_box(slide, "Set-LabEnvironment.ps1\nDiscovers resources", 3.45, 2.1, 2.25, 0.9, fill="soft_teal", bold=True)
add_box(slide, "Environment variables\nDOCUMENTDB_CLUSTER_NAME\nAZURE_OPENAI_ENDPOINT\nAZURE_OPENAI_EMBEDDING_DEPLOYMENT", 6.25, 1.85, 2.75, 1.4, fill="soft_green", size=11, bold=True)
add_box(slide, "Notebooks\nLab 4 + Lab 5", 9.65, 2.1, 1.8, 0.9, fill="soft_amber", bold=True)
add_box(slide, "Azure services\nDocumentDB + Azure OpenAI", 5.0, 4.55, 3.4, 0.9, fill="soft_blue", bold=True)
for x1, x2 in [(2.95, 3.45), (5.7, 6.25), (9.0, 9.65)]:
    add_arrow(slide, x1, 2.55, x2, 2.55, "teal")
add_arrow(slide, 10.55, 3.0, 7.2, 4.55, "teal")
add_arrow(slide, 6.7, 3.25, 6.7, 4.55, "green")
add_text(slide, "No notebook hardcoding", 0.95, 3.55, 2.8, 0.35, size=16, bold=True, color="navy")
add_bullets(slide, ["Resource discovery belongs in setup", "Notebook code reads runtime contract", "AAD tokens replace copied secrets"], 0.95, 3.95, 3.2, 1.1, size=13)

# Slide 4
slide = new_slide("Lab 4 architecture: search over operational documents", "The executable path is vector search; BM25 and hybrid are explained as preview-gated extensions.", "Lab 4", "Start the Lab 4 story with why search needs multiple retrieval signals. Vector search is semantic similarity over embeddings. Full-text search, BM25, fuzzy, and phrase search are lexical signals. Hybrid combines those rankings. In this environment, if full-text search is unavailable, the lab should still run vector search and use the BM25/hybrid cells as design review material.")
add_box(slide, "Sample docs\nJSON content", 0.75, 2.0, 1.8, 0.8, fill="soft_blue", bold=True)
add_box(slide, "Embedding model\ntextembedding3small", 3.0, 2.0, 2.2, 0.8, fill="soft_teal", bold=True)
add_box(slide, "DocumentDB collection\nworkshop_content", 5.65, 2.0, 2.25, 0.8, fill="soft_green", bold=True)
add_box(slide, "Vector index\ncosmosSearch", 8.35, 1.55, 1.9, 0.75, fill="soft_amber", bold=True)
add_box(slide, "Full-text index\nBM25 preview", 8.35, 2.55, 1.9, 0.75, fill="soft_red", bold=True)
add_box(slide, "Queries\nvector | BM25 | hybrid", 10.75, 2.0, 1.9, 0.8, fill="soft_blue", bold=True)
for x1, x2 in [(2.55, 3.0), (5.2, 5.65), (7.9, 8.35), (10.25, 10.75)]:
    add_arrow(slide, x1, 2.4, x2, 2.4, "teal")
add_bullets(slide, [
    "The collection is the integration point: raw fields plus embedding vectors.",
    "Vector search can be demonstrated as the dependable path.",
    "BM25, fuzzy, phrase, and hybrid should be narrated when full-text preview returns code 115 or equivalent capability errors.",
], 1.0, 4.25, 11.2, 1.6, size=16)
add_timeline(slide, ["Setup", "Load docs", "Embed", "Index", "Vector demo", "Preview search talk-through"], y=6.65)

# Slide 5
slide = new_slide("Lab 4 opening talk track", "First five minutes: remove ambiguity before code runs.", "Lab 4", "Use this as the natural opening. Say: We are going to build search as a layered retrieval system. The notebook already knows the cluster and endpoint through environment variables. If the setup cell fails, it means the runtime contract is missing, not that the notebook needs hardcoded values. For BM25 and hybrid, today we may not execute them if the preview gate is closed. That is fine. We will read the query shape and discuss what changes when the capability is enabled.")
add_bullets(slide, [
    "Position the lab: semantic retrieval first, lexical retrieval second, fusion last.",
    "State the contract: Set-LabEnvironment.ps1 populates names; notebooks should not be edited per student.",
    "Call out identity: AzureCliCredential obtains tokens for DocumentDB and Cognitive Services.",
    "Set expectation: unavailable BM25/hybrid is not a lab failure; it is a capability boundary.",
], 0.9, 2.0, 11.7, 2.4, size=18)
add_box(slide, "Say this plainly", 0.95, 5.0, 2.1, 0.5, fill="soft_teal", bold=True)
add_text(slide, "If full-text search is gated, we will not fake it. We will execute the vector path and inspect the BM25/hybrid query contracts as production design material.", 3.35, 4.94, 8.6, 0.75, size=18, color="navy")

# Slide 6
slide = new_slide("Lab 4 technical flow", "What each notebook block is proving.", "Lab 4", "Walk the cells as tests. The connect cell proves identity and endpoint wiring. The load cell proves document shape. The embedding cell proves model access and vector dimensions. The vector index cell proves the retrieval index exists. The vector query proves a semantic signal. Full-text and hybrid cells prove the intended query contract, even if execution is disabled by service capability.")
add_box(slide, "1\nConnect", 0.75, 2.0, 1.25, 0.85, fill="soft_blue", bold=True)
add_box(slide, "2\nLoad docs", 2.35, 2.0, 1.25, 0.85, fill="soft_teal", bold=True)
add_box(slide, "3\nGenerate embeddings", 3.95, 2.0, 1.65, 0.85, fill="soft_green", bold=True)
add_box(slide, "4\nCreate vector index", 5.95, 2.0, 1.75, 0.85, fill="soft_amber", bold=True)
add_box(slide, "5\nRun vector query", 8.05, 2.0, 1.65, 0.85, fill="soft_blue", bold=True)
add_box(slide, "6\nTalk through BM25/hybrid", 10.05, 2.0, 2.35, 0.85, fill="soft_red", bold=True)
for x in [2.0, 3.6, 5.6, 7.7, 9.7]:
    add_arrow(slide, x, 2.43, x + 0.35, 2.43, "teal")
add_bullets(slide, [
    "Connect: token scopes and env variables are correct.",
    "Load: documents preserve fields used by lexical search later.",
    "Embed: query and document vectors must come from the same model.",
    "Index: vector index shape is the latency/recall tradeoff point.",
    "Preview search: BM25/hybrid are query contracts when feature access is blocked.",
], 1.1, 3.65, 11.3, 1.8, size=15)

# Slide 7
slide = new_slide("When BM25 is unavailable", "Turn a blocked query into an engineering explanation.", "Lab 4", "This is the most important facilitation slide for your constraint. Do not apologize for the feature boundary. Explain it. BM25 needs a full-text search index and service capability. The notebook should catch unsupported errors and keep going. Then read the aggregate pipeline: the search stage, index name, path, query, limit, and projection. Ask the audience how they would test scoring quality once access is enabled.")
add_box(slide, "Service capability\nFull-text preview", 0.85, 2.0, 2.35, 0.9, fill="soft_red", bold=True)
add_box(slide, "Index contract\ntextSearch / analyzer / path", 3.7, 2.0, 2.55, 0.9, fill="soft_amber", bold=True)
add_box(slide, "Query contract\n$search text / phrase / fuzzy", 6.75, 2.0, 2.55, 0.9, fill="soft_teal", bold=True)
add_box(slide, "Fallback\nvector-only retrieval", 9.8, 2.0, 2.35, 0.9, fill="soft_green", bold=True)
for x1, x2 in [(3.2, 3.7), (6.25, 6.75), (9.3, 9.8)]:
    add_arrow(slide, x1, 2.45, x2, 2.45, "teal")
add_bullets(slide, [
    "BM25 is lexical relevance, not semantic similarity.",
    "If unavailable, show the intended index/query shape and continue.",
    "Use vector search output to discuss recall and ranking behavior.",
    "Production question: how would you compare BM25, vector, and fused results offline?",
], 1.1, 3.75, 11.0, 1.7, size=17)

# Slide 8
slide = new_slide("Hybrid search without executing hybrid", "Explain fusion as a ranking algorithm, not magic in the database.", "Lab 4", "Describe hybrid as two candidate lists and a fusion function. One list comes from vector similarity, one from BM25. Reciprocal Rank Fusion rewards documents that appear high in either or both lists. Even if the hybrid operator is not runnable, attendees can understand the ranking mechanics and why hybrid often improves robustness for mixed semantic and keyword-heavy queries.")
add_box(slide, "Vector list\nsemantic neighbors", 1.0, 2.0, 2.25, 0.9, fill="soft_teal", bold=True)
add_box(slide, "BM25 list\nlexical matches", 1.0, 3.35, 2.25, 0.9, fill="soft_amber", bold=True)
add_box(slide, "RRF\n1 / (k + rank)", 4.45, 2.67, 2.1, 0.9, fill="soft_green", bold=True)
add_box(slide, "Final ranking\nmerged context", 7.75, 2.67, 2.35, 0.9, fill="soft_blue", bold=True)
add_box(slide, "LLM-ready evidence\ntop chunks/docs", 10.85, 2.67, 1.65, 0.9, fill="soft_red", bold=True)
add_arrow(slide, 3.25, 2.45, 4.45, 3.0, "teal")
add_arrow(slide, 3.25, 3.8, 4.45, 3.15, "teal")
add_arrow(slide, 6.55, 3.12, 7.75, 3.12, "teal")
add_arrow(slide, 10.1, 3.12, 10.85, 3.12, "teal")
add_bullets(slide, [
    "Use the diagram when the hybrid cell cannot run.",
    "Point out that fusion can also be done in application code for controlled experiments.",
    "Discuss evaluation: labeled queries, click logs, recall@k, precision@k, and answer faithfulness.",
], 1.05, 5.15, 11.4, 1.0, size=15)

# Slide 9
slide = new_slide("Lab 4 close: what the room should retain", "Five-minute synthesis before switching to RAG.", "Lab 4", "Close Lab 4 by converting the lab into mental models. Vector search answers 'meaning-like-this'. BM25 answers 'terms-like-this'. Hybrid answers 'combine the strengths'. Capability detection is part of production readiness. The notebook should not be modified per attendee; it should read environment variables and fail with an actionable setup command.")
add_bullets(slide, [
    "Vector search is the dependable executable path in this delivery environment.",
    "BM25 and hybrid are still worth teaching because they explain production retrieval quality.",
    "The runtime contract is environment-driven: cluster, endpoint, deployment name.",
    "A good search lab ends with evaluation, not with a single query result.",
], 1.0, 2.0, 11.0, 2.1, size=19)
add_box(slide, "Bridge to Lab 5", 1.0, 5.05, 2.0, 0.55, fill="soft_teal", bold=True)
add_text(slide, "Search gives us ranked evidence. RAG turns that evidence into a grounded response pattern with explicit prompt construction.", 3.25, 4.95, 8.7, 0.75, size=18, color="navy")

# Slide 10
slide = new_slide("Lab 5 architecture: RAG over DocumentDB", "The database stores chunks and vectors; the application owns retrieval policy and prompt assembly.", "Lab 5", "Open Lab 5 by saying: RAG is not 'call the model with a database attached.' It is an application architecture. We chunk source content, embed chunks, store them with metadata, retrieve evidence for a question, assemble a prompt, and call a chat model. In this environment, we can still teach the retrieval and prompt architecture even if lexical/hybrid signals are unavailable.")
add_box(slide, "Source content", 0.75, 2.0, 1.55, 0.75, fill="soft_blue", bold=True)
add_box(slide, "Chunking\nmetadata", 2.75, 2.0, 1.55, 0.75, fill="soft_teal", bold=True)
add_box(slide, "Embeddings", 4.75, 2.0, 1.55, 0.75, fill="soft_green", bold=True)
add_box(slide, "DocumentDB\nrag_chunks", 6.75, 2.0, 1.75, 0.75, fill="soft_amber", bold=True)
add_box(slide, "Retriever\nvector-first", 9.0, 2.0, 1.55, 0.75, fill="soft_blue", bold=True)
add_box(slide, "Prompt\nanswer", 11.0, 2.0, 1.45, 0.75, fill="soft_green", bold=True)
for x1, x2 in [(2.3, 2.75), (4.3, 4.75), (6.3, 6.75), (8.5, 9.0), (10.55, 11.0)]:
    add_arrow(slide, x1, 2.38, x2, 2.38, "teal")
add_bullets(slide, [
    "DocumentDB is both operational store and retrieval store for this workshop pattern.",
    "Chunk metadata is how you explain source attribution and debugging.",
    "Retriever behavior is policy: top-k, filters, thresholds, and fallback strategy.",
    "Prompt construction should make evidence boundaries visible.",
], 1.0, 4.0, 11.2, 1.7, size=16)

# Slide 11
slide = new_slide("Lab 5 30-minute pacing", "Keep the room out of low-value typing and inside the architecture.", "Lab 5", "This is your Lab 5 operating plan. Spend three minutes recapping from Lab 4, seven minutes on chunking and embeddings, eight minutes on retrieval, seven minutes on prompt assembly, and five minutes on production concerns. If the environment cannot execute BM25 or hybrid, use vector retrieval as the live demo and explain where the second retrieval signal would plug in.")
add_timeline(slide, ["30-33\nBridge", "33-40\nChunks + embeddings", "40-48\nRetriever", "48-55\nPrompt", "55-60\nProduction review"], y=2.0)
add_bullets(slide, [
    "Do not spend more than one minute on package or auth issues in the room.",
    "Use printed query shapes and diagrams for preview-gated features.",
    "Make students name the evidence set before talking about the answer.",
    "Close with observability: failed auth, empty retrieval, poor ranking, and hallucination risk.",
], 1.0, 3.15, 11.0, 1.9, size=18)

# Slide 12
slide = new_slide("Chunking is the first retrieval decision", "RAG quality is shaped before the vector index exists.", "Lab 5", "Explain that chunking is a modeling decision. Too small and context loses meaning. Too large and retrieval becomes noisy and expensive. Metadata is not decorative. It lets you filter, cite, debug, and evaluate. In a technical audience, this is the point to discuss deterministic preprocessing and repeatable chunk IDs.")
add_box(slide, "Document", 0.95, 2.1, 1.55, 0.7, fill="soft_blue", bold=True)
add_arrow(slide, 2.5, 2.45, 3.15, 2.45, "teal")
for idx in range(4):
    add_box(slide, f"Chunk {idx + 1}\ntext + metadata", 3.15 + idx * 1.65, 1.9 + (idx % 2) * 0.75, 1.3, 0.7, fill="soft_teal", size=11, bold=True)
add_arrow(slide, 9.75, 2.45, 10.55, 2.45, "teal")
add_box(slide, "Embed + store\nrag_chunks", 10.55, 2.1, 1.75, 0.7, fill="soft_green", bold=True)
add_bullets(slide, [
    "Chunk size is a precision/recall tradeoff.",
    "Metadata enables filtering and attribution.",
    "Stable chunk IDs make repeated loads and evaluations sane.",
    "For the walkthrough, focus on why each field exists in the stored document.",
], 1.0, 4.2, 11.2, 1.5, size=17)

# Slide 13
slide = new_slide("Retriever flow when hybrid is unavailable", "Keep retrieval live with vector search, then explain the hybrid insertion point.", "Lab 5", "Use this slide to avoid getting trapped by a blocked hybrid query. Run vector retrieval if available. Then say: the hybrid branch would add lexical candidates and fuse them before prompt assembly. The rest of the RAG architecture remains the same. This is valuable because production systems often have feature flags or region-specific capability differences.")
add_box(slide, "User question", 0.8, 2.05, 1.8, 0.75, fill="soft_blue", bold=True)
add_box(slide, "Embed query", 3.1, 2.05, 1.65, 0.75, fill="soft_teal", bold=True)
add_box(slide, "Vector top-k", 5.25, 1.55, 1.7, 0.75, fill="soft_green", bold=True)
add_box(slide, "BM25 top-k\nif enabled", 5.25, 2.75, 1.7, 0.75, fill="soft_red", bold=True)
add_box(slide, "Merge / rank", 7.6, 2.05, 1.65, 0.75, fill="soft_amber", bold=True)
add_box(slide, "Prompt context", 9.75, 2.05, 1.75, 0.75, fill="soft_blue", bold=True)
add_box(slide, "Answer", 11.85, 2.05, 0.9, 0.75, fill="soft_green", bold=True)
add_arrow(slide, 2.6, 2.43, 3.1, 2.43, "teal")
add_arrow(slide, 4.75, 2.43, 5.25, 1.95, "teal")
add_arrow(slide, 4.75, 2.43, 5.25, 3.15, "red")
add_arrow(slide, 6.95, 1.95, 7.6, 2.43, "teal")
add_arrow(slide, 6.95, 3.15, 7.6, 2.43, "red")
add_arrow(slide, 9.25, 2.43, 9.75, 2.43, "teal")
add_arrow(slide, 11.5, 2.43, 11.85, 2.43, "teal")
add_bullets(slide, [
    "Live path: question embedding to vector top-k to prompt context.",
    "Talk-through path: lexical top-k joins before merge/rank when enabled.",
    "The prompt and answer stages do not care which retrieval signals produced the evidence.",
], 1.0, 4.55, 11.1, 1.2, size=17)

# Slide 14
slide = new_slide("Prompt assembly: evidence, instruction, question", "Make grounding visible to the audience.", "Lab 5", "Show prompt construction as a data structure, not prose pasted into a chat box. The high-value concept is separation: system instruction, retrieved evidence, user question, and answer constraints. This is where the search lab becomes an AI application lab. If retrieval is weak, prompt quality cannot fully compensate.")
add_box(slide, "System instruction\nrole + constraints", 0.9, 2.0, 2.3, 0.9, fill="soft_blue", bold=True)
add_box(slide, "Retrieved context\nranked chunks + source", 3.65, 2.0, 2.55, 0.9, fill="soft_green", bold=True)
add_box(slide, "User question", 6.65, 2.0, 1.8, 0.9, fill="soft_teal", bold=True)
add_box(slide, "Answer policy\nuse evidence or say unknown", 8.9, 2.0, 2.45, 0.9, fill="soft_amber", bold=True)
add_box(slide, "Model response", 5.3, 4.2, 2.75, 0.85, fill="soft_red", bold=True)
for x in [3.2, 6.2, 8.45]:
    add_arrow(slide, x, 2.45, x + 0.45, 2.45, "teal")
add_arrow(slide, 10.1, 2.9, 6.65, 4.2, "teal")
add_bullets(slide, [
    "Ask attendees to identify which part is data and which part is policy.",
    "Call out injection risk: retrieved text is untrusted input.",
    "Evaluation target: faithfulness to the supplied context, not model fluency.",
], 1.0, 5.65, 11.0, 0.85, size=15)

# Slide 15
slide = new_slide("Troubleshooting script for the room", "Fast triage when a live cell fails.", "Operations", "This is a speaker handout slide. Keep it visible during questions. If env vars are missing, run Set-LabEnvironment.ps1 and restart the kernel. If DocumentDB auth fails, check az login and role assignment. If embedding fails, check deployment name and token scope. If BM25/hybrid fails, explain preview gating and continue with vector-only retrieval.")
add_bullets(slide, [
    "Missing env vars: run Set-LabEnvironment.ps1, restart terminal or notebook kernel.",
    "Embedding deployment not found: setup should use deployment name textembedding3small, not model name text-embedding-3-small.",
    "OpenAI auth error: endpoint uses Cognitive Services token scope, https://cognitiveservices.azure.com/.default.",
    "DocumentDB auth error: confirm az login and DocumentDB Entra user/role mapping.",
    "BM25 or hybrid unavailable: treat the cell as a query-shape walkthrough and continue vector retrieval.",
], 0.9, 2.0, 11.6, 3.0, size=17)
add_box(slide, "Rule of thumb", 0.95, 5.7, 2.0, 0.5, fill="soft_teal", bold=True)
add_text(slide, "Do not ask students to edit cluster names or deployment names in notebooks. Fix setup once, then let notebooks consume the contract.", 3.25, 5.6, 8.7, 0.7, size=18, color="navy")

# Slide 16
slide = new_slide("Production design questions", "Use these to close with a senior-engineering conversation.", "Wrap", "End with a conversation, not a recap. Ask how they would measure retrieval quality, what metadata they would add, whether they would choose vector-only or hybrid retrieval, how they would handle feature availability across regions, and what they would log for production support. This makes the session valuable even when some preview queries cannot execute live.")
add_bullets(slide, [
    "What metadata would you store with each chunk to support filters, citations, and deletion?",
    "How would you compare vector-only, BM25-only, and hybrid retrieval on your domain data?",
    "Where should capability detection live: notebook, app service, deployment pipeline, or feature flag?",
    "What do you log for a bad answer: query, retrieved IDs, scores, prompt, model version, and latency?",
    "When would you keep retrieval in DocumentDB versus introduce a separate search service?",
], 0.9, 2.0, 11.7, 3.2, size=18)

# Slide 17
slide = new_slide("Handout close", "What to say if BM25/hybrid cannot run today.", "Wrap", "Use this as your final script. Say: Today you saw the executable vector path and the architectural placement of BM25 and hybrid. The unavailable cells are not wasted; they are contracts. Once full-text preview is enabled, the same data model and prompt flow can accept lexical candidates and fused ranking. The most transferable lesson is not the syntax of one operator. It is how to design a retrieval pipeline that is explicit, observable, and resilient to capability boundaries.")
add_text(slide, "Today we are not blocked by a missing preview feature. We are separating what must execute from what must be understood.", 1.0, 2.0, 11.0, 0.85, size=25, bold=True, color="navy")
add_bullets(slide, [
    "Execute: setup, DocumentDB connection, embeddings, vector retrieval, prompt assembly.",
    "Explain: BM25 scoring, fuzzy/phrase matching, hybrid fusion, evaluation strategy.",
    "Leave them with: contracts, failure modes, and production tuning questions.",
], 1.15, 3.55, 10.8, 1.5, size=19)

prs.save(OUT_PATH)
print(OUT_PATH)
