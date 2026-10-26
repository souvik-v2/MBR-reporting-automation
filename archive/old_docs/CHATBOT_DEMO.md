# 🤖 Chatbot Demo - Interactive Examples

## Example 1: Quick Run with MBR Chatbot

```
$ py mbr_chatbot.py

================================================================================
🤖 MBR CHATBOT - Interactive Management Business Review Assistant
================================================================================

I'll help you create your MBR deliverables step by step!
Type 'help' for available commands, 'quit' to exit

🤖 > full

🚀 Running Complete MBR Workflow...

📊 Extracting Graph Data...
  Looking in: sample_data/Graph

✅ Success! Extracted data for 6 projects:
  • CANOPYNEXT: Week LCW 4/6 - 4/20
  • CONSUMERWEB: Week LCW 4/6 - 4/20
  • FLEX: Week LCW 4/6 - 4/20
  • PHOENIX: Week LCW 4/6 - 4/20
  • SEOD: Week LCW 4/6 - 4/20
  • WPM: Week LCW 4/6 - 4/20

💾 Saved to: output\graph_data_summary.xlsx

--------------------------------------------------------------------------------

🎯 Finding 'Goals Achieved in May' slides...
  Looking in: sample_data/ppt-data

✓ V2Solutions_SEOD_May_2026_Report.pptx: Found 1 goal slide(s)

Total goal slides found: 1

✓ Excel report saved: output\goals_achieved_report.xlsx
✓ Instructions saved: output\ppt_consolidation_instructions.txt

✅ Success! Found 1 goal slide(s)

💾 Saved to:
  • output\goals_achieved_report.xlsx
  • output\ppt_consolidation_instructions.txt

--------------------------------------------------------------------------------

🎉 Workflow Complete! All outputs saved to: output

📈 Status:
  ✅ Generated Files:
    • graph_data_summary.xlsx
    • goals_achieved_report.xlsx
    • ppt_consolidation_instructions.txt

🤖 > quit

👋 Goodbye! Happy reporting!
```

---

## Example 2: Natural Language with MBR Assistant

```
$ py mbr_assistant.py

================================================================================
🤖 MBR Assistant - Your Intelligent MBR Helper
================================================================================

Hi! I'm your MBR Assistant. I can help you with:
  
  📊 Extracting graph data from Excel files
  🎯 Finding 'Goals Achieved' slides in presentations
  📝 Creating consolidated reports
  💡 Guiding you through PowerPoint consolidation

Just tell me what you need in plain English!

Examples:
  • "Extract the graph data"
  • "Find goals for June"
  • "Run everything"

Type 'quit' to exit.

💬 You: hi, can you help me extract the graph data?

🤖 Assistant: Extracting graph data from LT-SimplifiedMetrics files...

🤖 Assistant: ✅ Success! Extracted data for 6 projects: CANOPYNEXT, CONSUMERWEB, FLEX, PHOENIX, SEOD, WPM

🤖 Assistant: 📁 Saved to: output\graph_data_summary.xlsx

💬 You: great! now find the goals for May

🤖 Assistant: Searching for 'Goals Achieved in May' slides...

🤖 Assistant: ✅ Found 1 goal slide(s)!

🤖 Assistant: 📁 Created:
  • goals_achieved_report.xlsx
  • ppt_consolidation_instructions.txt

💬 You: what have you created so far?

🤖 Assistant: Here's what I've created:
  ✅ graph_data_summary.xlsx
  ✅ goals_achieved_report.xlsx
  ✅ ppt_consolidation_instructions.txt

📁 All files are in: output/

💬 You: how do I consolidate the PowerPoint slides?

🤖 Assistant: Here's your PowerPoint consolidation guide:

GOALS ACHIEVED IN MAY - CONSOLIDATION GUIDE
================================================================================

SUMMARY:
  Total goal slides found: 1

SLIDES TO EXTRACT:

1. File: V2Solutions_SEOD_May_2026_Report.pptx
   Slide: #3
   Title: Goals Achieved in May


MANUAL CONSOLIDATION STEPS:
================================================================================

1. Open the template PowerPoint:
   sample_data/final-data/06-01-2026.pptx

2. For each slide listed above:
   a. Open the source presentation file
   b. Navigate to the specified slide number
   c. Right-click the slide in the thumbnail pane
   d. Select 'Copy'
   e. Go back to the template presentation
   f. Right-click after slide 5 (or desired position)
   g. Select 'Keep Source Formatting' when pasting

3. Save the consolidated presentation as:
   output/consolidated_mbr_manual.pptx

💬 You: change the month to June

🤖 Assistant: ✅ Updated month to: June

💬 You: thanks, bye!

🤖 Assistant: Goodbye! Feel free to come back anytime you need help with MBR! 👋
```

---

## Example 3: Step-by-Step with Configuration

```
$ py mbr_chatbot.py

🤖 > config

⚙️  Current Configuration:
  graph_folder: sample_data/Graph
  ppt_folder: sample_data/ppt-data
  template_path: sample_data/final-data/06-01-2026.pptx
  output_folder: output
  month: May

Change configuration? (y/n): y

⚙️  Change Configuration
Press Enter to keep current value
  graph_folder [sample_data/Graph]: 
  ppt_folder [sample_data/ppt-data]: 
  template_path [sample_data/final-data/06-01-2026.pptx]: 
  output_folder [output]: 
  month [May]: June

✅ Configuration updated!

⚙️  Current Configuration:
  graph_folder: sample_data/Graph
  ppt_folder: sample_data/ppt-data
  template_path: sample_data/final-data/06-01-2026.pptx
  output_folder: output
  month: June

🤖 > extract

📊 Extracting Graph Data...
  Looking in: sample_data/Graph

✅ Success! Extracted data for 6 projects:
  [... project list ...]

💾 Saved to: output\graph_data_summary.xlsx

🤖 > goals

🎯 Finding 'Goals Achieved in June' slides...
  Looking in: sample_data/ppt-data

Total goal slides found: 0

⚠️  No goal slides found for June

🤖 > status

📈 Status:
  ✅ Generated Files:
    • graph_data_summary.xlsx

🤖 > quit

👋 Goodbye! Happy reporting!
```

---

## Chatbot Features Demonstrated

✅ **Natural Language** (Assistant)
- Understands conversational input
- Flexible phrasing
- Context-aware responses

✅ **Menu-Driven** (Chatbot)
- Simple numbered commands
- Clear options
- Predictable behavior

✅ **Configuration**
- Change month
- Modify folders
- Customize output

✅ **Status Tracking**
- See what's been created
- Track progress
- Review outputs

✅ **Help System**
- Built-in guidance
- Command explanations
- Usage examples

---

## Quick Reference

### Chatbot Commands
| Input | Action |
|-------|--------|
| `1` or `extract` | Extract graph data |
| `2` or `goals` | Find goals slides |
| `3` or `full` | Run everything |
| `4` or `config` | Configure settings |
| `5` or `status` | Show status |
| `6` or `ppt` | PPT instructions |
| `help` | Show help |
| `quit` | Exit |

### Assistant Phrases
| Say | It Will |
|-----|---------|
| "Extract data" | Extract graph data |
| "Find goals" | Find goals slides |
| "Run all" | Run full workflow |
| "What did you do?" | Show status |
| "Help with PPT" | Show PPT guide |
| "Change to June" | Update month |

---

**Try it yourself!** Both chatbots are ready to use. 🚀
