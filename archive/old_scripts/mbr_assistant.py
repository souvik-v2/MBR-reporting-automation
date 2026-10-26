"""
Intelligent MBR Assistant with Natural Language Processing.

This assistant understands natural language commands like:
- "Extract graph data for May"
- "Find goals achieved slides"
- "Create the full report"
- "Show me what you've done"
- "Help me consolidate the PowerPoint"

Usage:
    py mbr_assistant.py
"""
from __future__ import annotations
import re
from pathlib import Path
from typing import Optional

from graph_data_extractor import extract_all_graph_data, create_graph_data_dataframe
from pptx_goals_extractor import create_goals_report


class MBRAssistant:
    """Intelligent assistant with natural language understanding."""
    
    def __init__(self):
        self.config = {
            'graph_folder': 'sample_data/Graph',
            'ppt_folder': 'sample_data/ppt-data',
            'template_path': 'sample_data/final-data/06-01-2026.pptx',
            'output_folder': 'output',
            'month': 'March'
        }
        self.results = {}
        self.conversation_history = []
    
    def greet(self):
        """Greet the user."""
        print("\n" + "=" * 80)
        print("🤖 MBR Assistant - Your Intelligent MBR Helper")
        print("=" * 80)
        print("""
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
  • "What have you created?"
  • "Help me with PowerPoint"
  • "Change the month to June"

Type 'quit' to exit.
""")
    
    def understand_intent(self, user_input: str) -> dict:
        """Parse user input to understand intent."""
        user_input = user_input.lower().strip()
        
        intent = {
            'action': None,
            'parameters': {}
        }
        
        # Extract month if mentioned
        months = ['january', 'february', 'march', 'april', 'may', 'june',
                  'july', 'august', 'september', 'october', 'november', 'december']
        for month in months:
            if month in user_input:
                intent['parameters']['month'] = month.capitalize()
        
        # Determine action
        if any(word in user_input for word in ['quit', 'exit', 'bye', 'goodbye']):
            intent['action'] = 'quit'
        
        elif any(word in user_input for word in ['help', 'what can you do', 'commands']):
            intent['action'] = 'help'
        
        elif any(phrase in user_input for phrase in ['extract graph', 'get graph', 'graph data', 'metrics']):
            intent['action'] = 'extract_graph'
        
        elif any(phrase in user_input for phrase in ['find goals', 'goals achieved', 'extract goals', 'goal slides']):
            intent['action'] = 'extract_goals'
        
        elif any(phrase in user_input for phrase in ['run everything', 'do everything', 'full workflow', 'complete', 'run all']):
            intent['action'] = 'full_workflow'
        
        elif any(phrase in user_input for phrase in ['show status', 'what have you', 'what did you', 'show results', 'list files']):
            intent['action'] = 'show_status'
        
        elif any(phrase in user_input for phrase in ['ppt', 'powerpoint', 'consolidate', 'merge slides']):
            intent['action'] = 'show_ppt_guide'
        
        elif any(phrase in user_input for phrase in ['config', 'settings', 'change month', 'set month']):
            intent['action'] = 'config'
        
        else:
            intent['action'] = 'unknown'
        
        return intent
    
    def respond(self, message: str):
        """Send a message to user."""
        print(f"\n🤖 Assistant: {message}\n")
    
    def extract_graph_data(self, month: Optional[str] = None):
        """Extract graph data."""
        if month:
            self.config['month'] = month
            self.respond(f"Got it! Looking for {month} data...")
        
        self.respond("Extracting graph data from LT-SimplifiedMetrics files...")
        
        try:
            graph_data = extract_all_graph_data(self.config['graph_folder'])
            df = create_graph_data_dataframe(graph_data)
            
            output_path = Path(self.config['output_folder']) / 'graph_data_summary.xlsx'
            output_path.parent.mkdir(parents=True, exist_ok=True)
            df.to_excel(output_path, index=False)
            
            projects = ', '.join(df['project'].tolist())
            self.respond(f"✅ Success! Extracted data for {len(df)} projects: {projects}")
            self.respond(f"📁 Saved to: {output_path}")
            
            self.results['graph_data'] = str(output_path)
        except Exception as e:
            self.respond(f"❌ Oops! Something went wrong: {e}")
    
    def extract_goals(self, month: Optional[str] = None):
        """Extract goals from presentations."""
        if month:
            self.config['month'] = month
        
        self.respond(f"Searching for 'Goals Achieved in {self.config['month']}' slides...")
        
        try:
            goals_excel = Path(self.config['output_folder']) / 'goals_achieved_report.xlsx'
            ppt_instructions = Path(self.config['output_folder']) / 'ppt_consolidation_instructions.txt'
            
            result = create_goals_report(
                ppt_data_folder=self.config['ppt_folder'],
                output_excel=str(goals_excel),
                output_instructions=str(ppt_instructions),
                month=self.config['month']
            )
            
            self.respond(f"✅ Found {result['goals_found']} goal slide(s)!")
            self.respond(f"📁 Created:\n  • {goals_excel.name}\n  • {ppt_instructions.name}")
            
            self.results['goals_report'] = str(goals_excel)
            self.results['ppt_instructions'] = str(ppt_instructions)
        except Exception as e:
            self.respond(f"❌ Oops! Something went wrong: {e}")
    
    def run_full_workflow(self, month: Optional[str] = None):
        """Run complete workflow."""
        self.respond("Running the complete MBR workflow for you...")
        
        self.extract_graph_data(month)
        print()
        self.extract_goals(month)
        
        self.respond("🎉 All done! Check the status to see what I created.")
    
    def show_status(self):
        """Show current status."""
        if not self.results:
            self.respond("I haven't created anything yet. What would you like me to do?")
            return
        
        message = "Here's what I've created:\n"
        for key, path in self.results.items():
            if Path(path).exists():
                message += f"  ✅ {Path(path).name}\n"
        
        message += f"\n📁 All files are in: {self.config['output_folder']}/"
        self.respond(message)
    
    def show_ppt_guide(self):
        """Show PowerPoint consolidation guide."""
        instructions_path = Path(self.config['output_folder']) / 'ppt_consolidation_instructions.txt'
        
        if not instructions_path.exists():
            self.respond("I haven't extracted the goals yet. Would you like me to do that first?")
            if input("Run goals extraction? (yes/no): ").lower().startswith('y'):
                self.extract_goals()
            return
        
        content = instructions_path.read_text(encoding='utf-8')
        self.respond(f"Here's your PowerPoint consolidation guide:\n\n{content}")
    
    def show_help(self):
        """Show help information."""
        self.respond("""I understand these requests:

📊 Graph Data:
  • "Extract graph data"
  • "Get the metrics"
  • "Pull the graph data"

🎯 Goals Extraction:
  • "Find goals achieved slides"
  • "Extract goals for June"
  • "Get the goals"

🚀 Full Workflow:
  • "Run everything"
  • "Do the complete workflow"
  • "Process all data"

📋 Status & Results:
  • "What have you done?"
  • "Show me the results"
  • "List the files"

💡 PowerPoint Help:
  • "Help with PowerPoint"
  • "Show PPT instructions"
  • "How do I consolidate slides?"

⚙️ Configuration:
  • "Change month to June"
  • "Set month to May"

Just ask me in plain English - I'll figure it out! 😊""")
    
    def run(self):
        """Main assistant loop."""
        self.greet()
        
        while True:
            try:
                user_input = input("💬 You: ").strip()
                
                if not user_input:
                    continue
                
                self.conversation_history.append(('user', user_input))
                
                intent = self.understand_intent(user_input)
                
                if intent['action'] == 'quit':
                    self.respond("Goodbye! Feel free to come back anytime you need help with MBR! 👋")
                    break
                
                elif intent['action'] == 'help':
                    self.show_help()
                
                elif intent['action'] == 'extract_graph':
                    self.extract_graph_data(intent['parameters'].get('month'))
                
                elif intent['action'] == 'extract_goals':
                    self.extract_goals(intent['parameters'].get('month'))
                
                elif intent['action'] == 'full_workflow':
                    self.run_full_workflow(intent['parameters'].get('month'))
                
                elif intent['action'] == 'show_status':
                    self.show_status()
                
                elif intent['action'] == 'show_ppt_guide':
                    self.show_ppt_guide()
                
                elif intent['action'] == 'config':
                    month = intent['parameters'].get('month', self.config['month'])
                    self.config['month'] = month
                    self.respond(f"✅ Updated month to: {month}")
                
                else:
                    self.respond("Hmm, I'm not sure what you want me to do. Type 'help' to see what I can do!")
            
            except KeyboardInterrupt:
                print("\n")
                self.respond("Goodbye! 👋")
                break
            except Exception as e:
                self.respond(f"❌ Oops, I encountered an error: {e}")


if __name__ == "__main__":
    assistant = MBRAssistant()
    assistant.run()
