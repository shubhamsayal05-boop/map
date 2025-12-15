"""
AVLDrive HeatMap Manager - Standalone Application
A complete standalone tool for managing HeatMap evaluations with built-in database,
data editor, and export capabilities. No Excel dependency.
"""

import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv
import json
from datetime import datetime
from collections import defaultdict
import os

# Color definitions
RED_COLOR = "#FF0000"
GREEN_COLOR = "#00FF00"
YELLOW_COLOR = "#FFFF00"

class Database:
    """SQLite database manager for evaluation data"""
    
    def __init__(self, db_path="heatmap_data.db"):
        self.db_path = db_path
        self.conn = None
        self.init_database()
    
    def init_database(self):
        """Initialize database with required tables"""
        self.conn = sqlite3.connect(self.db_path)
        cursor = self.conn.cursor()
        
        # Operations table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS operations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                opcode TEXT UNIQUE NOT NULL,
                operation_name TEXT NOT NULL,
                is_parent INTEGER DEFAULT 0,
                parent_opcode TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Evaluations table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS evaluations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                opcode TEXT NOT NULL,
                operation_name TEXT NOT NULL,
                tested_avl REAL,
                driv_p1 TEXT,
                driv_target REAL,
                driv_tested REAL,
                driv_status TEXT,
                resp_p1 TEXT,
                resp_target REAL,
                resp_tested REAL,
                resp_status TEXT,
                final_status TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (opcode) REFERENCES operations(opcode)
            )
        ''')
        
        # HeatMap results table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS heatmap_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                opcode TEXT NOT NULL,
                operation_name TEXT NOT NULL,
                status TEXT,
                status_color TEXT,
                generated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (opcode) REFERENCES operations(opcode)
            )
        ''')
        
        self.conn.commit()
    
    def add_operation(self, opcode, operation_name, is_parent=False, parent_opcode=None):
        """Add a new operation"""
        cursor = self.conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO operations (opcode, operation_name, is_parent, parent_opcode)
                VALUES (?, ?, ?, ?)
            ''', (opcode, operation_name, 1 if is_parent else 0, parent_opcode))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
    
    def add_evaluation(self, data):
        """Add evaluation data"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO evaluations (
                opcode, operation_name, tested_avl, driv_p1, driv_target, 
                driv_tested, driv_status, resp_p1, resp_target, resp_tested,
                resp_status, final_status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            data['opcode'], data['operation'], data.get('tested_avl'),
            data.get('driv_p1'), data.get('driv_target'), data.get('driv_tested'),
            data.get('driv_status'), data.get('resp_p1'), data.get('resp_target'),
            data.get('resp_tested'), data.get('resp_status'), data.get('final_status')
        ))
        self.conn.commit()
    
    def get_all_evaluations(self):
        """Get all evaluation records"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM evaluations ORDER BY opcode')
        return cursor.fetchall()
    
    def get_evaluations_by_opcode(self, opcode):
        """Get evaluations for a specific opcode"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM evaluations WHERE opcode = ?', (opcode,))
        return cursor.fetchall()
    
    def get_all_operations(self):
        """Get all operations"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM operations ORDER BY opcode')
        return cursor.fetchall()
    
    def delete_evaluation(self, eval_id):
        """Delete an evaluation"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM evaluations WHERE id = ?', (eval_id,))
        self.conn.commit()
    
    def update_evaluation(self, eval_id, data):
        """Update evaluation data"""
        cursor = self.conn.cursor()
        cursor.execute('''
            UPDATE evaluations SET
                operation_name = ?, tested_avl = ?, driv_status = ?,
                resp_status = ?, final_status = ?
            WHERE id = ?
        ''', (
            data['operation'], data.get('tested_avl'),
            data.get('driv_status'), data.get('resp_status'),
            data.get('final_status'), eval_id
        ))
        self.conn.commit()
    
    def save_heatmap_result(self, opcode, operation_name, status, status_color):
        """Save generated heatmap result"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO heatmap_results (opcode, operation_name, status, status_color)
            VALUES (?, ?, ?, ?)
        ''', (opcode, operation_name, status, status_color))
        self.conn.commit()
    
    def get_latest_heatmap_results(self):
        """Get latest heatmap results"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT opcode, operation_name, status, status_color
            FROM heatmap_results
            WHERE generated_at = (SELECT MAX(generated_at) FROM heatmap_results)
            ORDER BY opcode
        ''')
        return cursor.fetchall()
    
    def clear_data(self, table_name):
        """Clear all data from a table"""
        cursor = self.conn.cursor()
        cursor.execute(f'DELETE FROM {table_name}')
        self.conn.commit()
    
    def close(self):
        """Close database connection"""
        if self.conn:
            self.conn.close()


class AIAnalyzer:
    """AI-powered analysis engine"""
    
    def analyze_evaluation_data(self, evaluations):
        """Analyze evaluation data for quality and completeness"""
        analysis = {
            'total_evaluations': len(evaluations),
            'status_distribution': defaultdict(int),
            'recommendations': [],
            'warnings': [],
            'quality_score': 100
        }
        
        # Count status distribution
        for eval_data in evaluations:
            final_status = eval_data[12] if len(eval_data) > 12 else None
            if final_status and final_status != 'N/A':
                analysis['status_distribution'][final_status] += 1
        
        # Calculate quality metrics
        red_count = analysis['status_distribution'].get('RED', 0)
        total_count = sum(analysis['status_distribution'].values())
        
        if total_count > 0:
            failure_rate = (red_count / total_count) * 100
            
            if failure_rate > 50:
                analysis['recommendations'].append(
                    f"⚠️ High failure rate detected ({failure_rate:.1f}%). "
                    "Consider reviewing test procedures or requirements."
                )
                analysis['quality_score'] -= 30
            elif failure_rate > 30:
                analysis['recommendations'].append(
                    f"⚡ Moderate failure rate ({failure_rate:.1f}%). "
                    "Some operations may need attention."
                )
                analysis['quality_score'] -= 15
            else:
                analysis['recommendations'].append(
                    f"✓ Good test performance ({100-failure_rate:.1f}% pass rate)."
                )
        
        if analysis['total_evaluations'] < 20:
            analysis['warnings'].append(
                "📊 Limited evaluation data. Consider testing more operations."
            )
            analysis['quality_score'] -= 10
        
        return analysis


class StandaloneHeatMapApp:
    """Main standalone application"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("AVLDrive HeatMap Manager - Standalone")
        self.root.geometry("1200x800")
        
        self.db = Database()
        self.ai_analyzer = AIAnalyzer()
        
        self.setup_ui()
        self.load_default_operations()
    
    def setup_ui(self):
        """Setup the user interface"""
        # Create menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Import from CSV", command=self.import_csv)
        file_menu.add_command(label="Export to CSV", command=self.export_csv)
        file_menu.add_command(label="Export to PDF", command=self.export_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Data menu
        data_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Data", menu=data_menu)
        data_menu.add_command(label="Clear All Evaluations", command=self.clear_evaluations)
        data_menu.add_command(label="Initialize Operations", command=self.load_default_operations)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Data Entry Tab
        self.create_data_entry_tab()
        
        # HeatMap Tab
        self.create_heatmap_tab()
        
        # AI Analysis Tab
        self.create_analysis_tab()
    
    def create_data_entry_tab(self):
        """Create data entry interface"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="📝 Data Entry")
        
        # Toolbar
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="➕ Add Evaluation", command=self.add_evaluation_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="✏️ Edit Selected", command=self.edit_evaluation_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🗑️ Delete Selected", command=self.delete_evaluation).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🔄 Refresh", command=self.refresh_data_grid).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="📋 Paste from Excel", command=self.paste_from_clipboard).pack(side=tk.LEFT, padx=2)
        
        # Data grid
        grid_frame = ttk.Frame(frame)
        grid_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create Treeview with scrollbars
        columns = ('ID', 'OpCode', 'Operation', 'Tested AVL', 'Driv Status', 'Resp Status', 'Final Status')
        self.data_tree = ttk.Treeview(grid_frame, columns=columns, show='headings', height=20)
        
        for col in columns:
            self.data_tree.heading(col, text=col)
            self.data_tree.column(col, width=100 if col != 'Operation' else 250)
        
        # Scrollbars
        vsb = ttk.Scrollbar(grid_frame, orient="vertical", command=self.data_tree.yview)
        hsb = ttk.Scrollbar(grid_frame, orient="horizontal", command=self.data_tree.xview)
        self.data_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.data_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        grid_frame.grid_rowconfigure(0, weight=1)
        grid_frame.grid_columnconfigure(0, weight=1)
        
        self.refresh_data_grid()
    
    def create_heatmap_tab(self):
        """Create HeatMap generation and view tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="🗺️ HeatMap")
        
        # Toolbar
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="▶ Generate HeatMap", command=self.generate_heatmap).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="📊 Export Results", command=self.export_heatmap).pack(side=tk.LEFT, padx=2)
        
        # Results display
        results_frame = ttk.LabelFrame(frame, text="HeatMap Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        columns = ('OpCode', 'Operation', 'Status', 'Type')
        self.heatmap_tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=25)
        
        for col in columns:
            self.heatmap_tree.heading(col, text=col)
            self.heatmap_tree.column(col, width=150 if col != 'Operation' else 350)
        
        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.heatmap_tree.yview)
        self.heatmap_tree.configure(yscrollcommand=vsb.set)
        
        self.heatmap_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Configure tags for colors
        self.heatmap_tree.tag_configure('red', foreground='red')
        self.heatmap_tree.tag_configure('yellow', foreground='orange')
        self.heatmap_tree.tag_configure('green', foreground='green')
    
    def create_analysis_tab(self):
        """Create AI analysis tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="🤖 AI Analysis")
        
        # Toolbar
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="🔍 Run Analysis", command=self.run_ai_analysis).pack(side=tk.LEFT, padx=2)
        
        # Results display
        self.analysis_text = scrolledtext.ScrolledText(frame, height=30, width=100)
        self.analysis_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    def load_default_operations(self):
        """Load default operation structure"""
        default_ops = [
            ("10000000", "AVL-DRIVE Rating", True, None),
            ("10100000", "Drive away", True, None),
            ("10101300", "Creep", False, "10100000"),
            ("10101100", "Standing start", False, "10100000"),
            ("10102400", "Rolling start", False, "10100000"),
            ("10120000", "Acceleration", True, None),
            ("10120100", "Full load", False, "10120000"),
            ("10120200", "Constant load", False, "10120000"),
            ("10120300", "Load increase", False, "10120000"),
            ("10120900", "Load decrease", False, "10120000"),
            ("10030000", "Tip in", True, None),
            ("10030100", "At deceleration", False, "10030000"),
            ("10030200", "At constant speed / acceleration", False, "10030000"),
            ("10040000", "Tip out", True, None),
            ("10040300", "At constant speed / acceleration", False, "10040000"),
            ("10040400", "At deceleration", False, "10040000"),
            ("10070000", "Deceleration", True, None),
            ("10070500", "Without brake", False, "10070000"),
            ("10070100", "Transition to constant speed", False, "10070000"),
            ("10071000", "Constant Brake", False, "10070000"),
            ("10090000", "Gear shift", True, None),
            ("10092300", "Power-on upshift", False, "10090000"),
            ("10092500", "Tip out upshift", False, "10090000"),
            ("10098400", "Load reversal upshift", False, "10090000"),
            ("10092100", "Coast / brake-on upshift", False, "10090000"),
            ("10093200", "Power-on downshift", False, "10090000"),
            ("10093100", "Kick down / tip in downshift", False, "10090000"),
            ("10093400", "Coast / brake-on downshift", False, "10090000"),
            ("10097800", "Maneuvering", False, "10090000"),
            ("10097900", "Selector lever change", False, "10090000"),
        ]
        
        for opcode, name, is_parent, parent in default_ops:
            self.db.add_operation(opcode, name, is_parent, parent)
        
        messagebox.showinfo("Success", "Default operations initialized!")
    
    def add_evaluation_dialog(self):
        """Show dialog to add new evaluation"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Evaluation")
        dialog.geometry("500x400")
        
        # Form fields
        ttk.Label(dialog, text="OpCode:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        opcode_entry = ttk.Entry(dialog, width=30)
        opcode_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="Operation:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        operation_entry = ttk.Entry(dialog, width=30)
        operation_entry.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="Tested AVL:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        avl_entry = ttk.Entry(dialog, width=30)
        avl_entry.grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="Driv Status:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        driv_status_combo = ttk.Combobox(dialog, values=['GREEN', 'YELLOW', 'RED', 'N/A'], width=28)
        driv_status_combo.grid(row=3, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="Resp Status:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        resp_status_combo = ttk.Combobox(dialog, values=['GREEN', 'YELLOW', 'RED', 'N/A'], width=28)
        resp_status_combo.grid(row=4, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="Final Status:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
        final_status_combo = ttk.Combobox(dialog, values=['GREEN', 'YELLOW', 'RED', 'N/A'], width=28)
        final_status_combo.grid(row=5, column=1, padx=5, pady=5)
        
        def save_evaluation():
            data = {
                'opcode': opcode_entry.get(),
                'operation': operation_entry.get(),
                'tested_avl': avl_entry.get() or None,
                'driv_status': driv_status_combo.get() or None,
                'resp_status': resp_status_combo.get() or None,
                'final_status': final_status_combo.get() or None
            }
            
            if not data['opcode'] or not data['operation']:
                messagebox.showerror("Error", "OpCode and Operation are required!")
                return
            
            self.db.add_evaluation(data)
            self.refresh_data_grid()
            dialog.destroy()
            messagebox.showinfo("Success", "Evaluation added!")
        
        ttk.Button(dialog, text="Save", command=save_evaluation).grid(row=6, column=0, columnspan=2, pady=20)
    
    def edit_evaluation_dialog(self):
        """Edit selected evaluation"""
        selected = self.data_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an evaluation to edit")
            return
        
        # Get selected item data
        item = self.data_tree.item(selected[0])
        values = item['values']
        eval_id = values[0]
        
        messagebox.showinfo("Edit", f"Edit dialog for evaluation ID {eval_id} (implementation pending)")
    
    def delete_evaluation(self):
        """Delete selected evaluation"""
        selected = self.data_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an evaluation to delete")
            return
        
        if messagebox.askyesno("Confirm", "Delete selected evaluation?"):
            item = self.data_tree.item(selected[0])
            eval_id = item['values'][0]
            self.db.delete_evaluation(eval_id)
            self.refresh_data_grid()
            messagebox.showinfo("Success", "Evaluation deleted!")
    
    def refresh_data_grid(self):
        """Refresh data grid with current evaluations"""
        # Clear existing items
        for item in self.data_tree.get_children():
            self.data_tree.delete(item)
        
        # Load evaluations
        evaluations = self.db.get_all_evaluations()
        for eval_data in evaluations:
            self.data_tree.insert('', 'end', values=(
                eval_data[0],  # ID
                eval_data[1],  # OpCode
                eval_data[2],  # Operation
                eval_data[3] or '',  # Tested AVL
                eval_data[7] or '',  # Driv Status
                eval_data[11] or '',  # Resp Status
                eval_data[12] or ''  # Final Status
            ))
    
    def paste_from_clipboard(self):
        """Paste data from clipboard (Excel-like)"""
        try:
            clipboard_data = self.root.clipboard_get()
            lines = clipboard_data.strip().split('\n')
            
            added_count = 0
            for line in lines:
                parts = line.split('\t')
                if len(parts) >= 2:
                    data = {
                        'opcode': parts[0].strip(),
                        'operation': parts[1].strip(),
                        'tested_avl': parts[2].strip() if len(parts) > 2 else None,
                        'driv_status': parts[3].strip() if len(parts) > 3 else None,
                        'resp_status': parts[4].strip() if len(parts) > 4 else None,
                        'final_status': parts[5].strip() if len(parts) > 5 else None
                    }
                    
                    if data['opcode'] and data['operation']:
                        self.db.add_evaluation(data)
                        added_count += 1
            
            self.refresh_data_grid()
            messagebox.showinfo("Success", f"Added {added_count} evaluations from clipboard!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste data: {str(e)}")
    
    def generate_heatmap(self):
        """Generate HeatMap from evaluations"""
        # Clear previous results
        for item in self.heatmap_tree.get_children():
            self.heatmap_tree.delete(item)
        
        # Get all evaluations
        evaluations = self.db.get_all_evaluations()
        eval_by_opcode = defaultdict(list)
        
        for eval_data in evaluations:
            opcode = eval_data[1]
            final_status = eval_data[12]
            if final_status and final_status != 'N/A':
                eval_by_opcode[opcode].append(final_status)
        
        # Get all operations
        operations = self.db.get_all_operations()
        parent_statuses = {}
        
        # Process sub-operations first
        for op in operations:
            opcode = op[1]
            op_name = op[2]
            is_parent = op[3]
            parent_opcode = op[4]
            
            if not is_parent and opcode in eval_by_opcode:
                # Sub-operation
                statuses = eval_by_opcode[opcode]
                worst_status = self.get_worst_status(statuses)
                
                # Track for parent calculation
                if parent_opcode:
                    if parent_opcode not in parent_statuses:
                        parent_statuses[parent_opcode] = []
                    parent_statuses[parent_opcode].append(worst_status)
                
                # Add to tree
                color_tag = worst_status.lower() if worst_status != 'N/A' else ''
                self.heatmap_tree.insert('', 'end', values=(
                    opcode, op_name, f"● ({worst_status})", "Sub-operation"
                ), tags=(color_tag,))
        
        # Process parent operations
        for op in operations:
            opcode = op[1]
            op_name = op[2]
            is_parent = op[3]
            
            if is_parent and opcode in parent_statuses:
                worst_status = self.get_worst_status(parent_statuses[opcode])
                status_text = self.get_text_for_status(worst_status)
                
                color_tag = worst_status.lower() if worst_status != 'N/A' else ''
                self.heatmap_tree.insert('', 'end', values=(
                    opcode, op_name, status_text, "Parent operation"
                ), tags=(color_tag,))
        
        messagebox.showinfo("Success", "HeatMap generated successfully!")
    
    def get_worst_status(self, statuses):
        """Get worst status from list"""
        if "RED" in statuses:
            return "RED"
        elif "YELLOW" in statuses:
            return "YELLOW"
        elif "GREEN" in statuses:
            return "GREEN"
        return "N/A"
    
    def get_text_for_status(self, status):
        """Get text for parent operation status"""
        if status == "RED":
            return "NOK"
        elif status == "GREEN":
            return "OK"
        elif status == "YELLOW":
            return "acceptable"
        return "N/A"
    
    def run_ai_analysis(self):
        """Run AI analysis on evaluation data"""
        self.analysis_text.delete(1.0, tk.END)
        
        evaluations = self.db.get_all_evaluations()
        analysis = self.ai_analyzer.analyze_evaluation_data(evaluations)
        
        self.analysis_text.insert(tk.END, "=== AI ANALYSIS RESULTS ===\n\n")
        self.analysis_text.insert(tk.END, f"Quality Score: {analysis['quality_score']}/100\n\n")
        
        self.analysis_text.insert(tk.END, f"Total Evaluations: {analysis['total_evaluations']}\n\n")
        
        self.analysis_text.insert(tk.END, "Status Distribution:\n")
        for status, count in analysis['status_distribution'].items():
            percentage = (count / sum(analysis['status_distribution'].values())) * 100 if analysis['status_distribution'] else 0
            self.analysis_text.insert(tk.END, f"  {status}: {count} ({percentage:.1f}%)\n")
        
        self.analysis_text.insert(tk.END, "\nRecommendations:\n")
        for rec in analysis['recommendations']:
            self.analysis_text.insert(tk.END, f"  {rec}\n")
        
        if analysis['warnings']:
            self.analysis_text.insert(tk.END, "\nWarnings:\n")
            for warn in analysis['warnings']:
                self.analysis_text.insert(tk.END, f"  {warn}\n")
    
    def clear_evaluations(self):
        """Clear all evaluation data"""
        if messagebox.askyesno("Confirm", "Clear all evaluations? This cannot be undone!"):
            self.db.clear_data('evaluations')
            self.refresh_data_grid()
            messagebox.showinfo("Success", "All evaluations cleared!")
    
    def import_csv(self):
        """Import data from CSV file"""
        filename = filedialog.askopenfilename(
            title="Import CSV",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'r') as f:
                    reader = csv.DictReader(f)
                    added_count = 0
                    for row in reader:
                        data = {
                            'opcode': row.get('OpCode', ''),
                            'operation': row.get('Operation', ''),
                            'tested_avl': row.get('Tested AVL'),
                            'driv_status': row.get('Driv Status'),
                            'resp_status': row.get('Resp Status'),
                            'final_status': row.get('Final Status')
                        }
                        if data['opcode'] and data['operation']:
                            self.db.add_evaluation(data)
                            added_count += 1
                
                self.refresh_data_grid()
                messagebox.showinfo("Success", f"Imported {added_count} evaluations!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import CSV: {str(e)}")
    
    def export_csv(self):
        """Export data to CSV file"""
        filename = filedialog.asksaveasfilename(
            title="Export CSV",
            filetypes=[("CSV Files", "*.csv")],
            defaultextension=".csv"
        )
        if filename:
            try:
                evaluations = self.db.get_all_evaluations()
                with open(filename, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['ID', 'OpCode', 'Operation', 'Tested AVL', 'Driv Status', 'Resp Status', 'Final Status'])
                    for eval_data in evaluations:
                        writer.writerow([
                            eval_data[0], eval_data[1], eval_data[2], eval_data[3] or '',
                            eval_data[7] or '', eval_data[11] or '', eval_data[12] or ''
                        ])
                
                messagebox.showinfo("Success", f"Exported to {filename}!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export CSV: {str(e)}")
    
    def export_pdf(self):
        """Export HeatMap to PDF"""
        messagebox.showinfo("PDF Export", "PDF export feature will be implemented with reportlab library")
    
    def export_heatmap(self):
        """Export HeatMap results"""
        filename = filedialog.asksaveasfilename(
            title="Export HeatMap",
            filetypes=[("CSV Files", "*.csv"), ("JSON Files", "*.json")],
            defaultextension=".csv"
        )
        if filename:
            try:
                items = []
                for item in self.heatmap_tree.get_children():
                    values = self.heatmap_tree.item(item)['values']
                    items.append({
                        'OpCode': values[0],
                        'Operation': values[1],
                        'Status': values[2],
                        'Type': values[3]
                    })
                
                if filename.endswith('.json'):
                    with open(filename, 'w') as f:
                        json.dump(items, f, indent=2)
                else:
                    with open(filename, 'w', newline='') as f:
                        writer = csv.DictWriter(f, fieldnames=['OpCode', 'Operation', 'Status', 'Type'])
                        writer.writeheader()
                        writer.writerows(items)
                
                messagebox.showinfo("Success", f"HeatMap exported to {filename}!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {str(e)}")


def main():
    """Main entry point"""
    root = tk.Tk()
    app = StandaloneHeatMapApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
