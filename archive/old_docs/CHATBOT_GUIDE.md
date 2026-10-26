# 🤖 MBR Chatbot & Assistant Guide

## Overview

We've created **two interactive tools** to make MBR workflow easier:

1. **MBR Chatbot** (`mbr_chatbot.py`) - Command-based interface
2. **MBR Assistant** (`mbr_assistant.py`) - Natural language interface

Both help you create MBR deliverables without remembering complex commands!

---

## 🤖 Option 1: MBR Chatbot (Command-Based)

### Start the Chatbot

```bash
py mbr_chatbot.py
```

### Available Commands

The chatbot understands simple numbered commands:

| Command | What It Does |
|---------|-------------|
| `1` or `extract` | Extract graph data from LT-SimplifiedMetrics files |
| `2` or `goals` | Find and extract "Goals Achieved" slides |
| `3` or `full` | Run complete workflow (extract + goals) |
| `4` or `config` | View/change configuration (folders, month, etc.) |
| `5` or `status` | Show what files have been created |
| `6` or `ppt` | Show PowerPoint consolidation instructions |
| `help` | Show all commands |
| `quit` | Exit chatbot |

### Example Session

```
🤖 MBR CHATBOT - Interactive Management Business Review Assistant
================================================================================

🤖 > help

📋 Available Commands:
  1. extract    - Extract graph data from LT-SimplifiedMetrics files
  2. goals      - Find and extract 'Goals Achieved' slides
  3. full       - Run complete workflow (extract + goals)
  ...

🤖 > 3

🚀 Running Complete MBR Workflow...

📊 Extracting Graph Data...
  Looking in: sample_data/Graph

✅ Success! Extracted data for 6 projects:
  • CANOPYNEXT: Week LCW 4/6 - 4/20
  • CONSUMERWEB: Week LCW 4/6 - 4/20
  ...

💾 Saved to: output\graph_data_summary.xlsx

🎯 Finding 'Goals Achieved in May' slides...

✅ Success! Found 1 goal slide(s)

🎉 Workflow Complete!

🤖 > status

📈 Status:
  ✅ Generated Files:
    • graph_data_summary.xlsx
    • goals_achieved_report.xlsx
    • ppt_consolidation_instructions.txt

🤖 > ppt

📋 PowerPoint Consolidation Instructions:
[Shows full instructions here]

🤖 > quit

👋 Goodbye! Happy reporting!
```

---

## 🧠 Option 2: MBR Assistant (Natural Language)

### Start the Assistant

```bash
py mbr_assistant.py
```

### Natural Language Understanding

The assistant understands **plain English commands**! No need to remember exact syntax.

### Example Commands

#### Extract Graph Data
```
💬 You: Extract the graph data
💬 You: Get me the metrics
💬 You: Pull the graph data for May
```

#### Find Goals
```
💬 You: Find goals achieved slides
💬 You: Extract goals for June
💬 You: Get the goals
```

#### Full Workflow
```
💬 You: Run everything
💬 You: Do the complete workflow
💬 You: Process all the data
```

#### Check Status
```
💬 You: What have you done?
💬 You: Show me the results
💬 You: List the files
```

#### PowerPoint Help
```
💬 You: Help with PowerPoint
💬 You: Show PPT instructions
💬 You: How do I consolidate slides?
```

#### Change Settings
```
💬 You: Change month to June
💬 You: Set month to May
```

### Example Session

```
🤖 MBR Assistant - Your Intelligent MBR Helper
================================================================================

Hi! I'm your MBR Assistant. I can help you with:
  📊 Extracting graph data from Excel files
  🎯 Finding 'Goals Achieved' slides in presentations
  📝 Creating consolidated reports
  💡 Guiding you through PowerPoint consolidation

Just tell me what you need in plain English!

💬 You: Run everything

🤖 Assistant: Running the complete MBR workflow for you...

🤖 Assistant: Extracting graph data from LT-SimplifiedMetrics files...

🤖 Assistant: ✅ Success! Extracted data for 6 projects: CANOPYNEXT, CONSUMERWEB, FLEX, PHOENIX, SEOD, WPM

🤖 Assistant: 📁 Saved to: output\graph_data_summary.xlsx

🤖 Assistant: Searching for 'Goals Achieved in May' slides...

🤖 Assistant: ✅ Found 1 goal slide(s)!

🤖 Assistant: 📁 Created:
  • goals_achieved_report.xlsx
  • ppt_consolidation_instructions.txt

🤖 Assistant: 🎉 All done! Check the status to see what I created.

💬 You: What did you create?

🤖 Assistant: Here's what I've created:
  ✅ graph_data_summary.xlsx
  ✅ goals_achieved_report.xlsx
  ✅ ppt_consolidation_instructions.txt

📁 All files are in: output/

💬 You: Help me with PowerPoint

🤖 Assistant: Here's your PowerPoint consolidation guide:

GOALS ACHIEVED IN MAY - CONSOLIDATION GUIDE
========================================
[Full instructions displayed]

💬 You: Change the month to June

🤖 Assistant: ✅ Updated month to: June

💬 You: bye

🤖 Assistant: Goodbye! Feel free to come back anytime you need help with MBR! 👋
```

---

## 🎯 Which One Should I Use?

### Use **MBR Chatbot** if you prefer:
- ✅ Simple numbered commands
- ✅ Clear menu-driven interface
- ✅ Traditional command-line style
- ✅ Precise control

### Use **MBR Assistant** if you prefer:
- ✅ Natural language conversations
- ✅ Talking in plain English
- ✅ More flexible interaction
- ✅ Conversational style

**Both do the same thing - choose what feels more comfortable!**

---

## 📊 What They Can Do

Both tools can:

1. **Extract Graph Data**
   - Reads `LT-SimplifiedMetrics-{project}.xlsx` files
   - Extracts V2/LT Carryover, Rework, Reopened SP
   - Saves to `output/graph_data_summary.xlsx`

2. **Find Goals Slides**
   - Scans PowerPoint presentations
   - Finds "Goals Achieved in {Month}" slides
   - Extracts content to `output/goals_achieved_report.xlsx`
   - Generates `output/ppt_consolidation_instructions.txt`

3. **Show Status**
   - Lists all generated files
   - Shows what's been completed

4. **Guide PowerPoint Consolidation**
   - Displays step-by-step instructions
   - Shows which slides to copy
   - Explains manual consolidation process

5. **Configure Settings**
   - Change month (May, June, etc.)
   - Change data folders
   - Customize output location

---

## 🚀 Quick Comparison

| Feature | Chatbot | Assistant |
|---------|---------|-----------|
| Commands | `1`, `extract`, `goals` | "Extract the data", "Find goals" |
| Interface | Menu-driven | Conversational |
| Ease of Use | Simple | Very simple |
| Learning Curve | None | None |
| Flexibility | Medium | High |

---

## 💡 Tips

1. **Try both** and see which you prefer!
2. Use `help` command to see what's available
3. Check `status` often to track progress
4. The `ppt` command shows consolidation instructions
5. Both save files to the same `output/` folder

---

## 🎓 Examples of Natural Language Commands

The **MBR Assistant** understands many variations:

```
✅ "Extract graph data"
✅ "Get the metrics"
✅ "Pull the data"
✅ "Find goals"
✅ "Extract goals for June"
✅ "Run everything"
✅ "Do it all"
✅ "What have you done?"
✅ "Show results"
✅ "Help with PowerPoint"
✅ "PPT instructions"
✅ "Change month to June"
✅ "Set to May"
```

---

## 📞 Getting Help

Within the chatbot/assistant:
- Type `help` to see available commands
- Type `status` to see what's been created
- Type `ppt` to see PowerPoint instructions

---

## 🎉 Benefits

- ❌ **No complex commands to remember**
- ✅ **Interactive and guided**
- ✅ **Immediate feedback**
- ✅ **Plain English (assistant)**
- ✅ **Step-by-step workflow**
- ✅ **Same powerful features as command-line**

---

**Enjoy your new interactive MBR tools!** 🤖✨
