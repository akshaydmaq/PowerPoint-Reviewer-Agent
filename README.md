# PowerPoint Reviewer Agent

An AI-powered agent that automatically reviews and corrects PowerPoint presentations using OpenAI GPT-4.

## Features

- **Spelling & Grammar Correction** - AI-powered detection and fixing of errors
- **Alignment Standardization** - Automatically aligns titles and elements consistently
- **Context-Aware Decisions** - Uses GPT-4 to understand context and preserve technical terms
- **Autonomous Agent Loop** - Iteratively analyzes and corrects all slides

## Architecture

| Component | Description |
|-----------|-------------|
| **LLM Brain** | GPT-4o for intelligent reasoning |
| **Agent Loop** | Autonomous iteration (max 20 cycles) |
| **State Management** | Tracks slides, corrections, completion |
| **Tool Calling** | 6 tools the agent can invoke |

### Available Tools

| Tool | Purpose |
|------|---------|
| `extract_slide_content` | Parse all text from PPTX |
| `analyze_text_for_errors` | AI-powered spell/grammar check |
| `analyze_alignment` | Detect misaligned elements |
| `add_correction` | Queue a fix |
| `apply_all_corrections` | Save corrected file |
| `mark_complete` | End the task |

## Installation

```bash
# Clone the repository
git clone https://github.com/akshaydmaq/PowerPoint-Reviewer-Agent.git
cd PowerPoint-Reviewer-Agent

# Install dependencies
pip install python-pptx openai python-dotenv
```

## Configuration

1. Create a `.env` file in the project root:
```
OPENAI_API_KEY=sk-your-api-key-here
```

2. Or set the environment variable:
```powershell
$env:OPENAI_API_KEY = "sk-your-api-key-here"
```

## Usage

```bash
# Run the agent
python pptx_agent.py
```

By default, the agent processes `Test deck.pptx` and outputs `Test deck_corrected.pptx`.

To process a different file, modify the `input_path` in `main()`.

## How It Works

1. **Extract** - Agent extracts all text content from the presentation
2. **Analyze** - Each text block is sent to GPT-4 for error detection
3. **Queue** - Corrections are added to a pending list
4. **Apply** - All corrections are applied to the PPTX file
5. **Save** - Corrected presentation is saved with `_corrected` suffix

## Files

| File | Description |
|------|-------------|
| `pptx_agent.py` | Main AI agent with OpenAI integration |
| `correct_pptx.py` | Simple rule-based corrector (no AI) |
| `analyze_pptx.py` | Utility to analyze PPTX content |

## Requirements

- Python 3.10+
- OpenAI API key
- Dependencies: `python-pptx`, `openai`, `python-dotenv`

## License

MIT
