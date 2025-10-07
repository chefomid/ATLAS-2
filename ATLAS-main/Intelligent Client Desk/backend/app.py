import os, time
from typing import List, Optional, Dict, Any
from fastapi import FastAPI
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware

# ---- Providers ---------------------------------------------------------------
class LLMProvider:
    def chat(self, messages: List[Dict[str, str]], **kw) -> str:
        raise NotImplementedError

class OpenAIProvider(LLMProvider):
    def __init__(self):
        from openai import OpenAI
        self.client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
        self.model = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")

    def chat(self, messages: List[Dict[str, str]], **kw) -> str:
        resp = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=kw.get("temperature", 0.2),
            max_tokens=kw.get("max_tokens", 800),
        )
        return resp.choices[0].message.content

# If you later get a Copilot-like API, implement it here.
class CopilotProvider(LLMProvider):
    def chat(self, messages: List[Dict[str, str]], **kw) -> str:
        # Placeholder: no public Copilot chat API available.
        raise RuntimeError("GitHub Copilot chat API is not publicly available.")

def get_provider() -> LLMProvider:
    provider = os.environ.get("IID_PROVIDER", "openai").lower()
    if provider == "openai":
        return OpenAIProvider()
    if provider == "copilot":
        return CopilotProvider()
    raise ValueError(f"Unknown IID_PROVIDER={provider}")

# ---- FastAPI app -------------------------------------------------------------
app = FastAPI(title="IID Backend", version="0.1.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True,
                   allow_methods=["*"], allow_headers=["*"])

class AskRequest(BaseModel):
    client: str
    question: str
    scope: Optional[str] = None
    snippets: Optional[List[str]] = None  # pass small text chunks you retrieved
    meta: Optional[Dict[str, Any]] = None

class AskResponse(BaseModel):
    answer: str
    latency_ms: int
    used_provider: str

@app.post("/ask", response_model=AskResponse)
def ask(req: AskRequest):
    """
    Minimal RAG-friendly shape:
    - You do retrieval outside (in the desktop app or another service)
    - Send only small snippets. Provider answers over those + question.
    """
    start = time.time()
    provider = get_provider()

    system = (
        "You are an insurance account analyst. "
        "Answer using only the provided snippets and scope when present. "
        "Cite snippet #s like [S1], [S2]. If unsure, say so."
    )
    messages = [{"role": "system", "content": system}]
    context = ""
    if req.scope:
        context += f"Scope: {req.scope}\n"
    if req.snippets:
        joined = "\n\n".join(f"[S{i+1}] {s}" for i, s in enumerate(req.snippets))
        context += f"Snippets:\n{joined}\n"
    if context:
        messages.append({"role": "user", "content": context})
    messages.append({"role": "user", "content": f"Client: {req.client}\nQuestion: {req.question}"})

    answer = provider.chat(messages)
    return AskResponse(
        answer=answer,
        latency_ms=int((time.time() - start) * 1000),
        used_provider=provider.__class__.__name__,
    )
