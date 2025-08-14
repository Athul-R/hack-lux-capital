"""
Modal.ai backend for Google Sheets AI Assistant (coding-focused)

This file is adapted from the provided markdown. It defines:
- A Modal app with model/session volumes
- A CodingContextManager for session context
- A function to (optionally) download coding models
- A coding_llm_inference function (stubbed, replace with real llama.cpp call)
- A web endpoint (coding-query) that the Chrome extension calls

To deploy:
  pip install modal
  modal token new
  modal deploy google-sheets-llm/modal_server.py

After deploy, update the Chrome extension background endpoint to your URL:
  https://<your-app-name>--coding-query.modal.run
"""

from __future__ import annotations

import json
import uuid
from pathlib import Path
from typing import Dict, List, Any

import modal


# Modal configuration
MINUTES = 60
app = modal.App("google-sheets-coding-llm")

# Simple base image with Python 3.12 and FastAPI
image = (
    modal.Image.debian_slim(python_version="3.12")
    .pip_install("fastapi==0.111.0", "uvicorn==0.30.1")
)

# Model and session storage (volumes)
model_cache = modal.Volume.from_name("coding-llm-cache", create_if_missing=True)
cache_dir = "/root/.cache/models"

session_volume = modal.Volume.from_name("coding-llm-sessions", create_if_missing=True)
session_dir = "/root/sessions"


@app.function(image=image, volumes={cache_dir: model_cache}, timeout=30 * MINUTES)
def download_coding_models() -> None:
    """Placeholder to download coding models.

    Replace with actual Hugging Face downloads if desired.
    This function exists to mirror the markdown's structure.
    """
    print("No-op model download. Customize download_coding_models() as needed.")
    model_cache.commit()


class CodingContextManager:
    """Manage conversation context per session, with a coding system prompt."""

    def __init__(self, session_dir_path: str):
        self.session_dir = Path(session_dir_path)

    def load_session(self, session_id: str) -> List[Dict[str, Any]]:
        session_file = self.session_dir / f"{session_id}.json"
        if session_file.exists():
            try:
                return json.loads(session_file.read_text())
            except Exception:
                return []
        return []

    def save_session(self, session_id: str, messages: List[Dict[str, Any]]) -> None:
        try:
            self.session_dir.mkdir(parents=True, exist_ok=True)
            session_file = self.session_dir / f"{session_id}.json"
            session_file.write_text(json.dumps(messages, indent=2))
        except Exception as e:
            print(f"Error saving session {session_id}: {e}")

    def add_message(
        self,
        session_id: str,
        role: str,
        content: str,
        metadata: Dict[str, Any] | None = None,
    ) -> List[Dict[str, Any]]:
        messages = self.load_session(session_id)

        # Inject system prompt on first message
        if not messages:
            system_prompt = self._create_coding_system_prompt(metadata)
            messages.append({"role": "system", "content": system_prompt})

        messages.append({"role": role, "content": content})

        # Simple truncation if too long (placeholder for summarization)
        max_messages = 40
        if len(messages) > max_messages:
            messages = [messages[0]] + messages[-(max_messages - 1) :]

        self.save_session(session_id, messages)
        return messages

    def _create_coding_system_prompt(self, metadata: Dict[str, Any] | None) -> str:
        return (
            "You are an expert coding assistant specialized in Excel/Google Sheets automation and programming tasks.\n\n"
            f"Current spreadsheet context:\n{json.dumps(metadata, indent=2) if metadata else 'No spreadsheet data available'}\n\n"
            "Your expertise includes:\n"
            "- Excel formula generation (VLOOKUP, INDEX/MATCH, SUMIF, PIVOT, etc.)\n"
            "- VBA/Visual Basic programming\n"
            "- Google Apps Script\n"
            "- Data analysis and manipulation\n"
            "- Financial and investment banking calculations\n"
            "- JavaScript for web automation\n\n"
            "Always provide:\n"
            "1. Clear, working code solutions\n"
            "2. Step-by-step explanations\n"
            "3. Specific cell references and ranges\n"
            "4. Copy-paste ready formulas\n"
            "5. Error handling considerations\n\n"
            "Focus on practical, production-ready solutions that integrate seamlessly with spreadsheet environments."
        )


@app.function(
    image=image,
    volumes={cache_dir: model_cache, session_dir: session_volume},
    timeout=15 * MINUTES,
)
def coding_llm_inference(
    session_id: str,
    prompt: str,
    metadata: Dict[str, Any] | None = None,
    model_choice: str = "phi-3.5-mini",
) -> Dict[str, Any]:
    """Run coding-focused inference. Stubbed for now.

    Replace this with a call to llama.cpp or a hosted LLM, as per the markdown.
    """
    ctx_manager = CodingContextManager(session_dir)

    # Build a context-augmented prompt (mirrors markdown intent)
    coding_prompt = (
        f"Your expertise includes:
            - Excel formula generation (VLOOKUP, INDEX/MATCH, SUMIF, etc.)
            - VBA/Visual Basic programming
            - Google Apps Script
            - Data analysis and manipulation
            - Financial and investment banking calculations
            - Code generation in multiple programming languages\n\n"
        f"Spreadsheet Context: {json.dumps(metadata, indent=2) if metadata else 'No spreadsheet loaded'}\n\n"
        f"User Request: {prompt}\n\n"
        "Please provide a practical solution with:\n"
        "1. Working code/formulas\n"
        "2. Clear explanations\n"
        "3. Integration steps for Google Sheets/Excel\n"
        "4. Specific cell references where applicable\n"
        "5. Error handling considerations\n\n"
        "Focus on practical, production-ready code that integrates well with spreadsheet environments."
    )

    messages = ctx_manager.add_message(session_id, "user", coding_prompt, metadata)

    # Stubbed LLM response
    response_text = (
        "Here is a suggested approach based on your request.\n\n"
        "Example formula: `=SUM(A:A)`\n"
        "Example Apps Script snippet:\n"
        "```javascript\nfunction example() {\n  Logger.log('Hello from AI Assistant');\n}\n```\n"
    )

    ctx_manager.add_message(session_id, "assistant", response_text, metadata)
    session_volume.commit()

    return {
        "session_id": session_id,
        "response": response_text,
        "model_used": model_choice,
        "metadata": metadata or {},
    }


@app.function(image=image, volumes={session_dir: session_volume})
@modal.web_endpoint(method="POST", label="coding-query")
def coding_query_endpoint(
    session_id: str | None = None,
    prompt: str = "",
    metadata: dict | None = None,
    model: str = "phi-3.5-mini",
):
    """Web endpoint for the Chrome extension."""
    if not session_id:
        session_id = str(uuid.uuid4())
    if not prompt:
        return {"error": "No prompt provided"}

    result = coding_llm_inference.remote(session_id, prompt, metadata or {}, model)
    return result


@app.local_entrypoint()
def setup_coding_models():
    """Optional: download/setup models (no-op by default)."""
    download_coding_models.remote()
    print("✅ Model setup complete (no-op). Customize to download GGUF models.")


