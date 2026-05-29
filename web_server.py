import uvicorn
from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from typing import List, Dict, Any

from llm_agent import chat as agent_chat

app = FastAPI(title="SAP Web Chat")

_HTML = """<!DOCTYPE html>
<html lang="zh">
<head>
  <meta charset="UTF-8">
  <title>SAP 助手</title>
  <style>
    body { font-family: sans-serif; max-width: 800px; margin: 40px auto; padding: 0 20px; }
    #messages { height: 420px; overflow-y: auto; border: 1px solid #ddd; padding: 12px;
                margin-bottom: 10px; border-radius: 4px; background: #fafafa; }
    .user { text-align: right; color: #0066cc; margin: 6px 0; }
    .assistant { text-align: left; color: #333; margin: 6px 0; white-space: pre-wrap; }
    #input-row { display: flex; gap: 8px; }
    #msg { flex: 1; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
    button { padding: 8px 18px; cursor: pointer; border-radius: 4px;
             background: #0066cc; color: white; border: none; }
    button:hover { background: #0052a3; }
  </style>
</head>
<body>
  <h2>SAP 助手</h2>
  <div id="messages"></div>
  <div id="input-row">
    <input id="msg" placeholder="输入指令，例如：查询订单 1000001"
           onkeydown="if(event.key==='Enter')send()">
    <button onclick="send()">发送</button>
  </div>
  <script>
    let history = [];
    async function send() {
      const input = document.getElementById('msg');
      const text = input.value.trim();
      if (!text) return;
      input.value = '';
      append('user', text);
      const res = await fetch('/chat', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({message: text, history})
      });
      const data = await res.json();
      history = data.history;
      append('assistant', data.reply);
    }
    function append(role, text) {
      const div = document.createElement('div');
      div.className = role;
      div.textContent = (role === 'user' ? '你：' : 'SAP 助手：') + text;
      document.getElementById('messages').appendChild(div);
      document.getElementById('messages').scrollTop = 9999;
    }
  </script>
</body>
</html>"""


class ChatRequest(BaseModel):
    message: str
    history: List[Dict[str, Any]] = []


class ChatResponse(BaseModel):
    reply: str
    history: List[Dict[str, Any]]


@app.get("/", response_class=HTMLResponse)
async def index():
    return _HTML


@app.post("/chat", response_model=ChatResponse)
async def chat(req: ChatRequest):
    reply, updated_history = agent_chat(req.message, req.history)
    return ChatResponse(reply=reply, history=updated_history)


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
