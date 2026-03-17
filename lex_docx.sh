#!/usr/bin/env bash
# lex_docx — DOCX 自动化工具库 CLI wrapper
# Persisted at /srcfile/scripts/lex_docx; symlinked to /usr/local/bin/lex_docx by start.sh.
# Works from any directory inside the container.

LEX_DOCX_DIR="/root/.openclaw/workspace-lex/tools"
export PYTHONPATH="${LEX_DOCX_DIR}${PYTHONPATH:+:$PYTHONPATH}"
exec python3 -m lex_docx "$@"
