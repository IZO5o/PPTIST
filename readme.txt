Ai_ppt_generation 跑通说明（本地联调 + 远端 Ollama）
日期：2025-12-25

一、项目结构（简述）
- 后端：项目根目录的 main.py（FastAPI + LangChain），默认以 Ollama 作为 LLM Provider。
- 前端：PPTist/（Vite + Vue3），通过 /api 代理请求后端接口。

二、前置条件
1) 本机需要：Node.js + npm（用于启动前端）
2) 本机需要：Python 环境（用于启动后端）
3) 需要可用的 Ollama 服务（二选一）：
   - 方案A：Ollama 跑在后端同机（最简单）
   - 方案B：Ollama 跑在远端服务器（本机/后端机通过网络或 SSH 转发访问）

重要说明：
- 本项目后端当前代码是“固定使用 Ollama”，没有通过配置切换 OpenAI 等 Provider 的完整实现。
- 请勿在文档/IM 中传播 SSH 密码等敏感信息。

三、后端启动（FastAPI + uvicorn）
目的：启动后端 API（默认端口 8000），为前端提供 /tools/aippt_outline、/tools/aippt 等接口。

方式 1：直接运行 Python 入口
1) 进入项目根目录：
   cd /Users/zhengzhan/MyProject/Ai_ppt_generation
2) 启动：
   python main.py

方式 2：使用 uvicorn 命令启动
1) 进入项目根目录：
   cd /Users/zhengzhan/MyProject/Ai_ppt_generation
2) 启动：
   uvicorn main:app --host 0.0.0.0 --port 8000 --reload

验证：
- 打开接口文档：http://127.0.0.1:8000/docs

四、前端启动（PPTist / Vite）
目的：启动前端编辑器（默认端口 5173），通过前端 UI 调用后端 AI 接口。

关键点：必须在 PPTist/ 目录运行 npm 命令。

1) 进入前端目录（注意不是 PPTist/src）：
   cd /Users/zhengzhan/MyProject/Ai_ppt_generation/PPTist

2) 安装依赖（本次遇到 husky 权限导致 npm install 失败，使用 HUSKY=0 绕过）：
   HUSKY=0 npm install

   说明：
   - 报错现象：
     sh: node_modules/.bin/husky: Permission denied
     npm error code 126
   - 原因：husky prepare 脚本缺少可执行权限。
   - HUSKY=0 的作用：跳过 husky 的安装步骤，保证依赖能装好并启动前端。

3) 启动前端：
   npm run dev

验证：
- 打开：http://127.0.0.1:5173/

五、Ollama 配置（后端必须能访问 Ollama 的 11434）
目的：让后端能调用 LLM 生成大纲/内容。

后端默认读取环境变量：
- OLLAMA_BASE_URL（默认 http://localhost:11434）
- DEFAULT_MODEL（默认 deepseek-R1:latest）

这些变量从 config.py 读取，并支持 .env 文件（dotenv.load_dotenv）。

A. 远端 Ollama（推荐用于服务器部署）
1) 在“运行后端的机器”上测试远端是否可达：
   curl http://<OLLAMA服务器IP>:11434/api/tags

2) 若可达：设置环境变量并启动后端（同一个终端会话中）：
   export OLLAMA_BASE_URL="http://<OLLAMA服务器IP>:11434"
   python main.py

B. 远端 11434 不对外开放（用 SSH 端口转发）
场景：远端 Ollama 只允许本机访问或防火墙不放行 11434。

1) 在本机建立 SSH 转发（示例端口 4092 仅作为说明，以实际为准）：
   ssh -p <ssh端口> -L 11434:127.0.0.1:11434 root@<服务器IP>

2) 此时本机会出现一个本地端口：
   http://127.0.0.1:11434  -> 实际转发到远端 Ollama

3) 在另一个终端启动后端并指向本地转发地址：
   export OLLAMA_BASE_URL="http://127.0.0.1:11434"
   python main.py

六、模型选择与常见报错处理
1) 404（模型不存在）
现象：
- 后端日志：Ollama call failed with status code 404
原因：
- 前端传给后端的 model 名称在 Ollama /api/tags 列表中不存在。
处理：
- 将前端模型选项/默认值改为 Ollama 已存在的模型名。

2) 500（GPU 显存不足 / cudaMalloc OOM）
现象：
- 直接 curl /api/chat 返回：
  {"error":"... cudaMalloc failed: out of memory ..."}
原因：
- 运行 Ollama 的那台机器 GPU 显存不足以加载该模型。
处理：
- 优先选择更小模型（例如 gemma3:4b）。
- 或更换更大显存 GPU 机器 / 调整 Ollama 部署策略（CPU/量化/并发等）。

本次已验证：
- gemma3:4b 能正常响应 /api/chat
- qwen3:latest 与 qwen3:4b 在当前环境可能触发 CUDA OOM（取决于机器显存与当前负载）

七、本次为跑通做过的前端改动（模型名）
目的：避免 404/500，让前端默认选择可用模型。

修改点：
- PPTist/src/views/Editor/AIPPTDialog.vue
  - 模型下拉选项改为 Ollama /api/tags 中存在的模型
  - 默认模型切换为 gemma3:4b
- PPTist/src/services/index.ts
  - AI_Writing 默认模型切换为 gemma3:4b

八、快速自检清单
1) 后端是否起来：
- http://127.0.0.1:8000/docs 能打开

2) 前端是否起来：
- http://127.0.0.1:5173 能打开

3) Ollama 是否可用：
- 本机（或转发后）：
  curl http://127.0.0.1:11434/api/tags
- 最小对话测试：
  curl http://127.0.0.1:11434/api/chat \
    -H 'Content-Type: application/json' \
    -d '{"model":"gemma3:4b","messages":[{"role":"user","content":"hello"}],"stream":false}'

九、生产/内网部署概念（简述）
“部署到公司内网”通常指：
- 服务运行在公司内网服务器/内网 K8s 上
- 只允许公司网络（或 VPN）访问，不对公网暴露
- 端口开放与访问控制由防火墙/安全组/网关统一管理

建议部署形态：
- 后端与 Ollama 同机（最简单）或同内网可达
- 前端使用 build 后的静态文件通过 nginx 提供（而不是 npm run dev）
