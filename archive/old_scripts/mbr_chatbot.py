"""
Interactive MBR Chatbot - User-friendly interface for MBR workflow.

This chatbot guides users through:
1. Extracting graph data
2. Finding goal slides
3. Creating consolidated reports
4. Step-by-step PowerPoint consolidation

Usage:
    py mbr_chatbot.py
"""
from __future__ import annotations
import sys
from pathlib import Path
from typing import Dict, Any

from graph_data_extractor import extract_all_graph_data, create_graph_data_dataframe
from pptx_goals_extractor import create_goals_report


class MBRChatbot:
    """Interactive chatbot for MBR workflow."""
    
    def __init__(self):
        self.config = {
            'graph_folder': 'sample_data/Graph',
            'ppt_folder': 'sample_data/ppt-data',
            'template_path': 'sample_data/final-data/06-01-2026.pptx',
            'output_folder': 'output',
            'month': 'March'
        }
        self.results = {}
    
    def print_header(self):
        """Print chatbot header."""
        print("\n" + "=" * 80)
        print("🤖 MBR CHATBOT - Interactive Management Business Review Assistant")
        print("=" * 80)
        print("\nI'll help you create your MBR deliverables step by step!")
        print("Type 'help' for available commands, 'quit' to exit\n")
    
    def print_help(self):
        """Print available commands."""
        print("\n📋 Available Commands:")
        print("  1. extract    - Extract graph data from LT-SimplifiedMetrics files")
        print("  2. goals      - Find and extract 'Goals Achieved' slides")
        print("  3. full       - Run complete workflow (extract + goals)")
        print("  4. config     - View/change configuration")
        print("  5. status     - Show current status and outputs")
        print("  6. ppt        - Show PowerPoint consolidation instructions")
        print("  7. help       - Show this help message")
        print("  8. quit       - Exit chatbot\n")
    
    def show_config(self):
        """Display current configuration."""
        print("\n⚙️  Current Configuration:")
        for key, value in self.config.items():
            print(f"  {key}: {value}")
        print()
    
    def change_config(self):
        """Allow user to change configuration."""
        print("\n⚙️  Change Configuration")
        print("Press Enter to keep current value")
        
        for key in self.config.keys():
            current = self.config[key]
            new_value = input(f"  {key} [{current}]: ").strip()
            if new_value:
                self.config[key] = new_value
        
        print("\n✅ Configuration updated!")
        self.show_config()
    
    def extract_graph_data(self):
        """Extract graph data with user feedback."""
        print("\n📊 Extracting Graph Data...")
        print(f"  Looking in: {self.config['graph_folder']}")
        
        try:
            graph_data = extract_all_graph_data(self.config['graph_folder'])
            df = create_graph_data_dataframe(graph_data)
            
            output_path = Path(self.config['output_folder']) / 'graph_data_summary.xlsx'
            output_path.parent.mkdir(parents=True, exist_ok=True)
            df.to_excel(output_path, index=False)
            
            print(f"\n✅ Success! Extracted data for {len(df)} projects:")
            for _, row in df.iterrows():
                print(f"  • {row['project']}: Week {row.get('week', 'N/A')}")
            
            print(f"\n💾 Saved to: {output_path}")
            self.results['graph_data'] = str(output_path)
            
            return True
        except Exception as e:
            print(f"\n❌ Error: {e}")
            return False
    
    def extract_goals(self):
        """Extract goals from presentations."""
        print(f"\n🎯 Finding 'Goals Achieved in {self.config['month']}' slides...")
        print(f"  Looking in: {self.config['ppt_folder']}")
        
        try:
            goals_excel = Path(self.config['output_folder']) / 'goals_achieved_report.xlsx'
            ppt_instructions = Path(self.config['output_folder']) / 'ppt_consolidation_instructions.txt'
            
            result = create_goals_report(
                ppt_data_folder=self.config['ppt_folder'],
                output_excel=str(goals_excel),
                output_instructions=str(ppt_instructions),
                month=self.config['month']
            )
            
            print(f"\n✅ Success! Found {result['goals_found']} goal slide(s)")
            print(f"\n💾 Saved to:")
            print(f"  • {goals_excel}")
            print(f"  • {ppt_instructions}")
            
            self.results['goals_report'] = str(goals_excel)
            self.results['ppt_instructions'] = str(ppt_instructions)
            
            return True
        except Exception as e:
            print(f"\n❌ Error: {e}")
            return False
    
    def run_full_workflow(self):
        """Run complete workflow."""
        print("\n🚀 Running Complete MBR Workflow...")
        
        success = True
        if self.extract_graph_data():
            print("\n" + "-" * 80)
        else:
            success = False
        
        if self.extract_goals():
            print("\n" + "-" * 80)
        else:
            success = False
        
        if success:
            print("\n🎉 Workflow Complete! All outputs saved to:", self.config['output_folder'])
            self.show_status()
    
    def show_status(self):
        """Show current status and outputs."""
        print("\n📈 Status:")
        
        if self.results:
            print("  ✅ Generated Files:")
            for key, path in self.results.items():
                if Path(path).exists():
                    print(f"    • {Path(path).name}")
        else:
            print("  ⚠️  No outputs generated yet. Run 'extract', 'goals', or 'full'")
        
        print()
    
    def show_ppt_instructions(self):
        """Show PowerPoint consolidation instructions."""
        instructions_path = Path(self.config['output_folder']) / 'ppt_consolidation_instructions.txt'
        
        if instructions_path.exists():
            print("\n📋 PowerPoint Consolidation Instructions:")
            print("-" * 80)
            content = instructions_path.read_text(encoding='utf-8')
            print(content)
        else:
            print("\n⚠️  Instructions not found. Run 'goals' or 'full' first.")
    
    def run(self):
        """Main chatbot loop."""
        self.print_header()
        
        while True:
            try:
                command = input("🤖 > ").strip().lower()
                
                if not command:
                    continue
                
                if command in ['quit', 'exit', 'q']:
                    print("\n👋 Goodbye! Happy reporting!")
                    break
                
                elif command in ['help', 'h', '?']:
                    self.print_help()
                
                elif command in ['1', 'extract', 'graph']:
                    self.extract_graph_data()
                
                elif command in ['2', 'goals']:
                    self.extract_goals()
                
                elif command in ['3', 'full', 'all']:
                    self.run_full_workflow()
                
                elif command in ['4', 'config', 'settings']:
                    self.show_config()
                    if input("\nChange configuration? (y/n): ").lower() == 'y':
                        self.change_config()
                
                elif command in ['5', 'status']:
                    self.show_status()
                
                elif command in ['6', 'ppt', 'instructions']:
                    self.show_ppt_instructions()
                
                else:
                    print(f"\n❓ Unknown command: '{command}'")
                    print("Type 'help' to see available commands")
            
            except KeyboardInterrupt:
                print("\n\n👋 Goodbye!")
                break
            except Exception as e:
                print(f"\n❌ Error: {e}")


if __name__ == "__main__":
    chatbot = MBRChatbot()
    chatbot.run()
