from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from openai import OpenAI
import json

# ----------------- 配置 -----------------
# 替换为你的阿里云 API KEY 和模型名称
API_KEY = "sk-00691987c8a2428a8b67459608c159c0"
MODEL_NAME = "qwen3-30b-a3b-instruct-2507"
API_BASE = "https://dashscope.aliyuncs.com/compatible-mode/v1" # OpenAI兼容接口地址

app = FastAPI()

# ----------------- 跨域资源共享 (CORS) -----------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ----------------- OpenAI 客户端初始化 -----------------
client = OpenAI(
    api_key=API_KEY,
    base_url=API_BASE
)

# ----------------- 请求体模型 -----------------
class DocumentRequest(BaseModel):
    documentContent: str

# ----------------- 接口路由 -----------------
@app.post("/api/fill-document")
async def fill_document(doc_request: DocumentRequest):
    """
    接收文档内容，调用AI模型提取信息，并返回填充数据。
    """
    # 自动识别文档中的待填充片段和内容，补全投标人、保证金、招标人、项目名称、日期等关键信息，返回 target/value 数组
    prompt = f"""
    你是一个专业的文档助理，请从以下“投标保证函”文件内容中自动识别所有需要填充的空白、占位符或待补全片段，并补全如下关键信息：
    - 投标人（如“投标人：邹晓东”）
    - 保证金金额（如“保证金 50万”）
    - 招标人（如“招标人：四川电信”）
    - 项目名称（如“项目名称：软件开发”）
    - 日期（如“日期：2025年8月9日”或文档中的签署日期）

    请返回一个 JSON 数组，每个元素包含：
    - target: 需要替换的原文片段（如“（投标人）：”或“__________”等）
    - value: 应填充的内容（如“邹晓东”、“50万”、“四川电信”、“软件开发”、“2025年8月9日”）

    “投标保证函”文件内容如下：
    \"\"\"{doc_request.documentContent}\"\"\"

    请只返回 JSON 数组，不要包含任何解释。例如：
    [
    {{"target": "（投标人）：", "value": "邹晓东"}},
    {{"target": "保证金", "value": "50万"}},
    {{"target": "（招标人）：", "value": "四川电信"}},
    {{"target": "项目名称", "value": "软件开发"}},
    {{"target": "日期", "value": "2025年8月9日"}}
    ]"""

    print("[fill-document] 输入文档内容:", doc_request.documentContent)
    print("[fill-document] 构造的Prompt:", prompt)
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "user", "content": prompt}
            ],
            response_format={ "type": "json_object" }
        )
        ai_response_content = response.choices[0].message.content
        print("[fill-document] AI原始响应:", ai_response_content)
        try:
            replacements = json.loads(ai_response_content)
            print("[fill-document] 解析后的replacements:", replacements)
        except json.JSONDecodeError:
            print("[fill-document] JSON解析失败，AI响应:", ai_response_content)
            raise HTTPException(status_code=500, detail="AI 返回的格式不正确")

        print("[fill-document] 返回结果:", {"success": True, "replacements": replacements})
        return {"success": True, "replacements": replacements}

    except Exception as e:
        print(f"[fill-document] Error calling AI API: {e}")
        raise HTTPException(status_code=500, detail=str(e))