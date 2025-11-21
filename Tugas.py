import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import networkx as nx
from collections import deque
from datetime import datetime, timedelta

class ProjectSchedulingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Project Scheduling Application - CPM & Gantt Chart")
        self.root.geometry("1400x800")
        self.root.configure(bg="#f0f0f0")
        
        # Data storage
        self.activities = []
        
        # Style configuration
        self.setup_styles()
        
        # Main container
        self.main_frame = tk.Frame(root, bg="#f0f0f0")
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create UI
        self.create_header()
        self.create_left_panel()
        self.create_right_panel()
        self.create_footer()
        
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Treeview style
        style.configure("Custom.Treeview",
                       background="#ffffff",
                       foreground="#000000",
                       rowheight=25,
                       fieldbackground="#ffffff",
                       borderwidth=0)
        style.map('Custom.Treeview', background=[('selected', '#4a90e2')])
        
        # Notebook style
        style.configure("TNotebook", background="#f0f0f0", borderwidth=0)
        style.configure("TNotebook.Tab", padding=[20, 10], font=('Arial', 10, 'bold'))
        
    def create_header(self):
        header_frame = tk.Frame(self.main_frame, bg="#2c3e50", height=80)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, 
                              text="📊 Project Scheduling & CPM Analysis",
                              font=('Arial', 24, 'bold'),
                              bg="#2c3e50",
                              fg="white")
        title_label.pack(pady=20)
        
    def create_left_panel(self):
        left_frame = tk.Frame(self.main_frame, bg="#ffffff", relief=tk.RAISED, bd=2)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 5))
        left_frame.configure(width=450)
        
        # Input section
        input_frame = tk.LabelFrame(left_frame, text="Input Kegiatan Baru", 
                                   font=('Arial', 12, 'bold'),
                                   bg="#ffffff", fg="#2c3e50", padx=10, pady=10)
        input_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Activity name
        tk.Label(input_frame, text="Nama Kegiatan:", bg="#ffffff", font=('Arial', 10)).grid(row=0, column=0, sticky='w', pady=5)
        self.activity_name = tk.Entry(input_frame, width=30, font=('Arial', 10))
        self.activity_name.grid(row=0, column=1, pady=5, padx=5)
        
        # Duration
        tk.Label(input_frame, text="Durasi (hari):", bg="#ffffff", font=('Arial', 10)).grid(row=1, column=0, sticky='w', pady=5)
        self.activity_duration = tk.Entry(input_frame, width=30, font=('Arial', 10))
        self.activity_duration.grid(row=1, column=1, pady=5, padx=5)
        
        # Dependencies
        tk.Label(input_frame, text="Dependensi (ex: 1,2,3):", bg="#ffffff", font=('Arial', 10)).grid(row=2, column=0, sticky='w', pady=5)
        self.activity_deps = tk.Entry(input_frame, width=30, font=('Arial', 10))
        self.activity_deps.grid(row=2, column=1, pady=5, padx=5)
        
        # Buttons
        button_frame = tk.Frame(input_frame, bg="#ffffff")
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        add_btn = tk.Button(button_frame, text="Tambah Kegiatan", 
                           command=self.add_activity,
                           bg="#27ae60", fg="white", 
                           font=('Arial', 10, 'bold'),
                           padx=10, pady=5, cursor="hand2")
        add_btn.pack(side=tk.LEFT, padx=5)
        
        clear_btn = tk.Button(button_frame, text="Clear All", 
                             command=self.clear_all,
                             bg="#e74c3c", fg="white",
                             font=('Arial', 10, 'bold'),
                             padx=10, pady=5, cursor="hand2")
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Import/Export buttons
        io_frame = tk.Frame(left_frame, bg="#ffffff")
        io_frame.pack(fill=tk.X, padx=10, pady=5)
        
        import_btn = tk.Button(io_frame, text="📥 Import Excel", 
                              command=self.import_excel,
                              bg="#3498db", fg="white",
                              font=('Arial', 10, 'bold'),
                              padx=10, pady=5, cursor="hand2")
        import_btn.pack(side=tk.LEFT, padx=5)
        
        export_btn = tk.Button(io_frame, text="📤 Export Excel", 
                              command=self.export_excel,
                              bg="#9b59b6", fg="white",
                              font=('Arial', 10, 'bold'),
                              padx=10, pady=5, cursor="hand2")
        export_btn.pack(side=tk.LEFT, padx=5)
        
        # Activities table
        table_frame = tk.LabelFrame(left_frame, text="Daftar Kegiatan", 
                                   font=('Arial', 12, 'bold'),
                                   bg="#ffffff", fg="#2c3e50")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollbar
        tree_scroll = ttk.Scrollbar(table_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        self.tree = ttk.Treeview(table_frame, 
                                yscrollcommand=tree_scroll.set,
                                style="Custom.Treeview",
                                selectmode='browse')
        self.tree.pack(fill=tk.BOTH, expand=True)
        tree_scroll.config(command=self.tree.yview)
        
        # Columns
        self.tree['columns'] = ('ID', 'Nama', 'Durasi', 'Dependensi')
        self.tree.column('#0', width=0, stretch=tk.NO)
        self.tree.column('ID', anchor=tk.CENTER, width=40)
        self.tree.column('Nama', anchor=tk.W, width=180)
        self.tree.column('Durasi', anchor=tk.CENTER, width=60)
        self.tree.column('Dependensi', anchor=tk.CENTER, width=120)
        
        # Headings
        self.tree.heading('ID', text='ID', anchor=tk.CENTER)
        self.tree.heading('Nama', text='Nama Kegiatan', anchor=tk.W)
        self.tree.heading('Durasi', text='Durasi', anchor=tk.CENTER)
        self.tree.heading('Dependensi', text='Dependensi', anchor=tk.CENTER)
        
        # Delete button
        delete_btn = tk.Button(left_frame, text="🗑️ Hapus Kegiatan Terpilih", 
                              command=self.delete_activity,
                              bg="#e67e22", fg="white",
                              font=('Arial', 10, 'bold'),
                              padx=10, pady=5, cursor="hand2")
        delete_btn.pack(pady=5)
        
    def create_right_panel(self):
        right_frame = tk.Frame(self.main_frame, bg="#ffffff", relief=tk.RAISED, bd=2)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # CPM Tab
        self.cpm_frame = tk.Frame(self.notebook, bg="#ffffff")
        self.notebook.add(self.cpm_frame, text="📊 CPM Analysis")
        
        # Network Diagram Tab
        self.network_frame = tk.Frame(self.notebook, bg="#ffffff")
        self.notebook.add(self.network_frame, text="🔗 Network Diagram")
        
        # Gantt Chart Tab
        self.gantt_frame = tk.Frame(self.notebook, bg="#ffffff")
        self.notebook.add(self.gantt_frame, text="📅 Gantt Chart")
        
        # Add scrollbar to each tab
        for frame in [self.cpm_frame, self.network_frame, self.gantt_frame]:
            canvas = tk.Canvas(frame, bg="#ffffff")
            scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg="#ffffff")
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e, c=canvas: c.configure(scrollregion=c.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Store references
            frame.canvas = canvas
            frame.scrollable_frame = scrollable_frame
            
    def create_footer(self):
        footer_frame = tk.Frame(self.main_frame, bg="#34495e", height=40)
        footer_frame.pack(fill=tk.X, pady=(10, 0))
        footer_frame.pack_propagate(False)
        
        copyright_label = tk.Label(footer_frame, 
                                  text="© 2025 Jevon Ivander Thomas - Project Management System",
                                  font=('Arial', 10),
                                  bg="#34495e",
                                  fg="white")
        copyright_label.pack(pady=10)
        
    def add_activity(self):
        name = self.activity_name.get().strip()
        duration = self.activity_duration.get().strip()
        deps = self.activity_deps.get().strip()
        
        if not name or not duration:
            messagebox.showwarning("Input Error", "Nama dan durasi harus diisi!")
            return
            
        try:
            duration = int(duration)
            if duration <= 0:
                raise ValueError()
        except:
            messagebox.showwarning("Input Error", "Durasi harus berupa angka positif!")
            return
            
        # Process dependencies
        dep_list = []
        if deps:
            try:
                dep_list = [int(x.strip()) for x in deps.split(',') if x.strip()]
            except:
                messagebox.showwarning("Input Error", "Format dependensi salah! Gunakan: 1,2,3")
                return
                
        activity_id = len(self.activities) + 1
        self.activities.append({
            'id': activity_id,
            'name': name,
            'duration': duration,
            'dependencies': dep_list
        })
        
        # Add to tree
        dep_str = ','.join(map(str, dep_list)) if dep_list else '-'
        self.tree.insert('', tk.END, values=(activity_id, name, duration, dep_str))
        
        # Clear inputs
        self.activity_name.delete(0, tk.END)
        self.activity_duration.delete(0, tk.END)
        self.activity_deps.delete(0, tk.END)
        
    def delete_activity(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Pilih kegiatan yang akan dihapus!")
            return
            
        item = self.tree.item(selected[0])
        activity_id = item['values'][0]
        
        # Remove from activities list
        self.activities = [a for a in self.activities if a['id'] != activity_id]
        
        # Re-index activities
        for idx, activity in enumerate(self.activities):
            activity['id'] = idx + 1
            
        # Refresh tree
        self.refresh_tree()
        
    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for activity in self.activities:
            dep_str = ','.join(map(str, activity['dependencies'])) if activity['dependencies'] else '-'
            self.tree.insert('', tk.END, values=(
                activity['id'],
                activity['name'],
                activity['duration'],
                dep_str
            ))
            
    def clear_all(self):
        if messagebox.askyesno("Konfirmasi", "Hapus semua kegiatan?"):
            self.activities = []
            self.refresh_tree()
            
    def import_excel(self):
        filename = filedialog.askopenfilename(
            title="Pilih File Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not filename:
            return
            
        try:
            df = pd.read_excel(filename)
            
            # Try different column name variations
            col_mappings = [
                {'name': ['nama', 'name', 'kegiatan', 'activity', 'task'],
                 'duration': ['durasi', 'duration', 'waktu', 'time'],
                 'deps': ['dependensi', 'dependencies', 'predecessor', 'prasyarat']},
            ]
            
            # Auto-detect columns
            df.columns = df.columns.str.lower().str.strip()
            
            name_col = None
            duration_col = None
            deps_col = None
            
            for col in df.columns:
                col_clean = col.lower().strip()
                if any(x in col_clean for x in col_mappings[0]['name']):
                    name_col = col
                elif any(x in col_clean for x in col_mappings[0]['duration']):
                    duration_col = col
                elif any(x in col_clean for x in col_mappings[0]['deps']):
                    deps_col = col
                    
            # If column names not found, try positional
            if not name_col and len(df.columns) >= 2:
                name_col = df.columns[0]
                duration_col = df.columns[1]
                if len(df.columns) >= 3:
                    deps_col = df.columns[2]
                    
            if not name_col or not duration_col:
                messagebox.showerror("Error", "Tidak dapat mendeteksi kolom nama/durasi!")
                return
                
            self.activities = []
            
            for idx, row in df.iterrows():
                name = str(row[name_col]).strip()
                if pd.isna(name) or name == '' or name.lower() == 'nan':
                    continue
                    
                try:
                    duration = int(float(row[duration_col]))
                except:
                    continue
                    
                deps = []
                if deps_col and not pd.isna(row[deps_col]):
                    deps_str = str(row[deps_col]).strip()
                    if deps_str and deps_str.lower() != 'nan':
                        try:
                            deps = [int(float(x.strip())) for x in deps_str.split(',') if x.strip()]
                        except:
                            pass
                            
                self.activities.append({
                    'id': len(self.activities) + 1,
                    'name': name,
                    'duration': duration,
                    'dependencies': deps
                })
                
            self.refresh_tree()
            messagebox.showinfo("Success", f"Berhasil mengimpor {len(self.activities)} kegiatan!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengimpor file: {str(e)}")
            
    def export_excel(self):
        if not self.activities:
            messagebox.showwarning("Warning", "Tidak ada data untuk diekspor!")
            return
            
        # Calculate CPM first
        cpm_result = self.calculate_cpm()
        if not cpm_result:
            messagebox.showwarning("Warning", "Gagal menghitung CPM!")
            return
            
        filename = filedialog.asksaveasfilename(
            title="Simpan File Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not filename:
            return
            
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # Activities sheet
                activities_df = pd.DataFrame([
                    {
                        'ID': a['id'],
                        'Nama Kegiatan': a['name'],
                        'Durasi': a['duration'],
                        'Dependensi': ','.join(map(str, a['dependencies'])) if a['dependencies'] else '-'
                    }
                    for a in self.activities
                ])
                activities_df.to_excel(writer, sheet_name='Activities', index=False)
                
                # CPM Results sheet
                cpm_df = pd.DataFrame([
                    {
                        'ID': k,
                        'Nama Kegiatan': v['name'],
                        'ES': v['ES'],
                        'EF': v['EF'],
                        'LS': v['LS'],
                        'LF': v['LF'],
                        'Slack': v['slack'],
                        'Jalur Kritis': 'Ya' if v['is_critical'] else 'Tidak'
                    }
                    for k, v in cpm_result.items()
                ])
                cpm_df.to_excel(writer, sheet_name='CPM Analysis', index=False)
                
            messagebox.showinfo("Success", "Data berhasil diekspor ke Excel!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengekspor file: {str(e)}")
            
    def calculate_cpm(self):
        if not self.activities:
            return None
            
        # Build activity dictionary
        act_dict = {a['id']: a for a in self.activities}
        
        # Calculate ES and EF (Forward pass)
        es = {a['id']: 0 for a in self.activities}
        ef = {}
        
        # Topological sort using Kahn's algorithm
        in_degree = {a['id']: len(a['dependencies']) for a in self.activities}
        queue = deque([a['id'] for a in self.activities if not a['dependencies']])
        
        processed = []
        while queue:
            current = queue.popleft()
            processed.append(current)
            
            # Calculate ES and EF
            if act_dict[current]['dependencies']:
                es[current] = max(ef[pred] for pred in act_dict[current]['dependencies'])
            ef[current] = es[current] + act_dict[current]['duration']
            
            # Update successors
            for activity in self.activities:
                if current in activity['dependencies']:
                    in_degree[activity['id']] -= 1
                    if in_degree[activity['id']] == 0:
                        queue.append(activity['id'])
                        
        if len(processed) != len(self.activities):
            messagebox.showerror("Error", "Terdapat circular dependency!")
            return None
            
        # Calculate LS and LF (Backward pass)
        project_duration = max(ef.values())
        lf = {a['id']: project_duration for a in self.activities if not any(a['id'] in act['dependencies'] for act in self.activities)}
        ls = {}
        
        # Process in reverse topological order
        for current in reversed(processed):
            if current not in lf:
                successors = [a['id'] for a in self.activities if current in a['dependencies']]
                if successors:
                    lf[current] = min(ls[succ] for succ in successors)
                else:
                    lf[current] = project_duration
                    
            ls[current] = lf[current] - act_dict[current]['duration']
            
        # Calculate slack and identify critical path
        result = {}
        for activity in self.activities:
            aid = activity['id']
            slack = ls[aid] - es[aid]
            result[aid] = {
                'name': activity['name'],
                'duration': activity['duration'],
                'ES': es[aid],
                'EF': ef[aid],
                'LS': ls[aid],
                'LF': lf[aid],
                'slack': slack,
                'is_critical': slack == 0,
                'dependencies': activity['dependencies']
            }
            
        return result
        
    def show_cpm_results(self):
        # Clear previous content
        for widget in self.cpm_frame.scrollable_frame.winfo_children():
            widget.destroy()
            
        cpm_result = self.calculate_cpm()
        if not cpm_result:
            return
            
        # Title
        title = tk.Label(self.cpm_frame.scrollable_frame, 
                        text="Critical Path Method (CPM) Analysis",
                        font=('Arial', 16, 'bold'),
                        bg="#ffffff", fg="#2c3e50")
        title.pack(pady=10)
        
        # Project duration
        duration_label = tk.Label(self.cpm_frame.scrollable_frame,
                                 text=f"Total Durasi Proyek: {max(v['EF'] for v in cpm_result.values())} hari",
                                 font=('Arial', 12, 'bold'),
                                 bg="#ffffff", fg="#27ae60")
        duration_label.pack(pady=5)
        
        # Critical path
        critical_activities = [v['name'] for v in cpm_result.values() if v['is_critical']]
        critical_label = tk.Label(self.cpm_frame.scrollable_frame,
                                 text=f"Jalur Kritis: {' → '.join(critical_activities)}",
                                 font=('Arial', 11),
                                 bg="#ffffff", fg="#e74c3c",
                                 wraplength=800)
        critical_label.pack(pady=5)
        
        # Results table
        table_frame = tk.Frame(self.cpm_frame.scrollable_frame, bg="#ffffff")
        table_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        # Create treeview
        columns = ('ID', 'Kegiatan', 'Durasi', 'ES', 'EF', 'LS', 'LF', 'Slack', 'Kritis')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # Column headings
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100 if col == 'Kegiatan' else 70, anchor=tk.CENTER)
            
        # Add data
        for aid, data in sorted(cpm_result.items()):
            values = (
                aid,
                data['name'],
                data['duration'],
                data['ES'],
                data['EF'],
                data['LS'],
                data['LF'],
                data['slack'],
                '✓' if data['is_critical'] else ''
            )
            
            tag = 'critical' if data['is_critical'] else 'normal'
            tree.insert('', tk.END, values=values, tags=(tag,))
            
        # Tag configuration
        tree.tag_configure('critical', background='#ffcccc')
        tree.tag_configure('normal', background='#ffffff')
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def zoom_factory(self, ax, base_scale=1.5):
        """Enable zoom with mouse wheel"""
        def zoom_fun(event):
            cur_xlim = ax.get_xlim()
            cur_ylim = ax.get_ylim()
            xdata = event.xdata
            ydata = event.ydata
            
            if xdata is None or ydata is None:
                return
                
            if event.button == 'up':
                # Zoom in
                scale_factor = 1 / base_scale
            elif event.button == 'down':
                # Zoom out
                scale_factor = base_scale
            else:
                return
                
            new_width = (cur_xlim[1] - cur_xlim[0]) * scale_factor
            new_height = (cur_ylim[1] - cur_ylim[0]) * scale_factor
            
            relx = (cur_xlim[1] - xdata) / (cur_xlim[1] - cur_xlim[0])
            rely = (cur_ylim[1] - ydata) / (cur_ylim[1] - cur_ylim[0])
            
            ax.set_xlim([xdata - new_width * (1 - relx), xdata + new_width * relx])
            ax.set_ylim([ydata - new_height * (1 - rely), ydata + new_height * rely])
            ax.figure.canvas.draw()
            
        fig = ax.get_figure()
        fig.canvas.mpl_connect('scroll_event', zoom_fun)
        
        return zoom_fun
    
    def pan_factory(self, ax):
        """Enable pan with left mouse button drag"""
        def on_press(event):
            if event.button != 1:  # Only left mouse button
                return
            ax._pan_start = (event.xdata, event.ydata)
            ax.figure.canvas.draw()
            
        def on_motion(event):
            if not hasattr(ax, '_pan_start') or ax._pan_start is None:
                return
            if event.button != 1:
                return
            if event.xdata is None or event.ydata is None:
                return
                
            dx = event.xdata - ax._pan_start[0]
            dy = event.ydata - ax._pan_start[1]
            
            cur_xlim = ax.get_xlim()
            cur_ylim = ax.get_ylim()
            
            ax.set_xlim([cur_xlim[0] - dx, cur_xlim[1] - dx])
            ax.set_ylim([cur_ylim[0] - dy, cur_ylim[1] - dy])
            
            ax.figure.canvas.draw()
            
        def on_release(event):
            if hasattr(ax, '_pan_start'):
                ax._pan_start = None
            ax.figure.canvas.draw()
            
        fig = ax.get_figure()
        fig.canvas.mpl_connect('button_press_event', on_press)
        fig.canvas.mpl_connect('motion_notify_event', on_motion)
        fig.canvas.mpl_connect('button_release_event', on_release)
        
    def show_network_diagram(self):
        # Clear previous content
        for widget in self.network_frame.scrollable_frame.winfo_children():
            widget.destroy()
            
        cpm_result = self.calculate_cpm()
        if not cpm_result:
            return
            
        # Create figure
        # Adjusted figsize to fit the screen better (approx 900x600 pixels)
        fig, ax = plt.subplots(figsize=(9, 6))
        fig.patch.set_facecolor('#ffffff')
        
        # Create directed graph
        G = nx.DiGraph()
        
        # Add nodes
        for aid, data in cpm_result.items():
            G.add_node(aid, label=f"{aid}\n{data['name']}\n({data['duration']}d)")
            
        # Add edges
        for aid, data in cpm_result.items():
            for dep in data['dependencies']:
                G.add_edge(dep, aid)
                
        # Layout
        try:
            pos = nx.spring_layout(G, k=2, iterations=50, seed=42)
        except:
            pos = nx.shell_layout(G)
            
        # Draw edges
        nx.draw_networkx_edges(G, pos, ax=ax, edge_color='#95a5a6', 
                              arrows=True, arrowsize=20, arrowstyle='->')
        
        # Draw critical path edges in red
        critical_edges = []
        for aid, data in cpm_result.items():
            if data['is_critical']:
                for dep in data['dependencies']:
                    if cpm_result[dep]['is_critical']:
                        critical_edges.append((dep, aid))
                        
        nx.draw_networkx_edges(G, pos, edgelist=critical_edges, ax=ax,
                              edge_color='#e74c3c', width=3,
                              arrows=True, arrowsize=20, arrowstyle='->')
        
        # Draw nodes
        node_colors = ['#e74c3c' if cpm_result[node]['is_critical'] else '#3498db' 
                      for node in G.nodes()]
        nx.draw_networkx_nodes(G, pos, ax=ax, node_color=node_colors,
                              node_size=3000, alpha=0.9)
        
        # Draw labels
        labels = {aid: f"{aid}\n{data['name'][:15]}\n{data['duration']}d" 
                 for aid, data in cpm_result.items()}
        nx.draw_networkx_labels(G, pos, labels, ax=ax, font_size=8,
                               font_weight='bold', font_color='white')
        
        # Use suptitle for better positioning
        fig.suptitle('Network Diagram\nKlik Kiri + Drag untuk Pan | Mouse Wheel untuk Zoom', 
                    fontsize=12, fontweight='bold', y=0.98)
        ax.axis('off')
        
        # Adjust layout to prevent title cutoff
        plt.tight_layout(rect=[0, 0, 1, 0.95])
        
        # Add Legend (Tkinter based)
        legend_frame = tk.Frame(self.network_frame.scrollable_frame, bg="#ffffff")
        legend_frame.pack(side=tk.TOP, pady=5)
        
        tk.Label(legend_frame, text="Keterangan Warna:", bg="#ffffff", font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        
        # Critical Path Legend
        lbl_crit = tk.Label(legend_frame, text=" Jalur Kritis ", bg="#e74c3c", fg="white", font=('Arial', 9, 'bold'))
        lbl_crit.pack(side=tk.LEFT, padx=5)
        
        # Normal Path Legend
        lbl_norm = tk.Label(legend_frame, text=" Kegiatan Normal ", bg="#3498db", fg="white", font=('Arial', 9, 'bold'))
        lbl_norm.pack(side=tk.LEFT, padx=5)
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.network_frame.scrollable_frame)
        ax.axis('off')
        
        # Adjust layout to prevent title cutoff
        plt.tight_layout(rect=[0, 0, 1, 0.95])
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.network_frame.scrollable_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add toolbar for other controls
        toolbar = NavigationToolbar2Tk(canvas, self.network_frame.scrollable_frame)
        toolbar.update()
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=10)
        
        # Enable mouse wheel zoom and pan
        self.zoom_factory(ax)
        self.pan_factory(ax)
        
    def show_gantt_chart(self):
        # Clear previous content
        for widget in self.gantt_frame.scrollable_frame.winfo_children():
            widget.destroy()
            
        cpm_result = self.calculate_cpm()
        if not cpm_result:
            return
            
        # Create figure
        # Adjusted width to fit screen, height dynamic but reasonable
        fig, ax = plt.subplots(figsize=(9, max(5, len(self.activities) * 0.5 + 2)))
        fig.patch.set_facecolor('#ffffff')
        
        # Sort activities by ID
        sorted_activities = sorted(cpm_result.items(), key=lambda x: x[0])
        
        # Y positions
        y_pos = range(len(sorted_activities))
        
        # Draw bars
        for idx, (aid, data) in enumerate(sorted_activities):
            color = '#e74c3c' if data['is_critical'] else '#3498db'
            
            # Main bar (ES to EF)
            ax.barh(idx, data['duration'], left=data['ES'], 
                   height=0.6, color=color, alpha=0.8,
                   edgecolor='black', linewidth=1.5)
            
            # Slack bar if exists
            if data['slack'] > 0:
                ax.barh(idx, data['slack'], left=data['EF'],
                       height=0.6, color='#95a5a6', alpha=0.3,
                       edgecolor='gray', linewidth=1)
                       
            # Add duration text
            mid_point = data['ES'] + data['duration'] / 2
            ax.text(mid_point, idx, f"{data['duration']}d",
                   ha='center', va='center', fontweight='bold',
                   color='white', fontsize=9)
            
        # Labels
        activity_labels = [f"{aid}. {data['name']}" for aid, data in sorted_activities]
        ax.set_yticks(y_pos)
        ax.set_yticklabels(activity_labels, fontsize=9)
        ax.set_xlabel('Hari', fontsize=11, fontweight='bold')
        
        # Use suptitle for better positioning
        fig.suptitle('Gantt Chart\nKlik Kiri + Drag untuk Pan | Mouse Wheel untuk Zoom', 
                    fontsize=12, fontweight='bold', y=0.98)
        
        # Grid
        ax.grid(axis='x', alpha=0.3, linestyle='--')
        ax.set_axisbelow(True)
        
        # Adjust layout to prevent title cutoff
        plt.tight_layout(rect=[0, 0, 1, 0.95])
        
        # Add Legend (Tkinter based)
        legend_frame = tk.Frame(self.gantt_frame.scrollable_frame, bg="#ffffff")
        legend_frame.pack(side=tk.TOP, pady=5)
        
        tk.Label(legend_frame, text="Keterangan Warna:", bg="#ffffff", font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        
        # Critical Path Legend
        lbl_crit = tk.Label(legend_frame, text=" Jalur Kritis ", bg="#e74c3c", fg="white", font=('Arial', 9, 'bold'))
        lbl_crit.pack(side=tk.LEFT, padx=5)
        
        # Normal Path Legend
        lbl_norm = tk.Label(legend_frame, text=" Kegiatan Normal ", bg="#3498db", fg="white", font=('Arial', 9, 'bold'))
        lbl_norm.pack(side=tk.LEFT, padx=5)
        
        # Slack Time Legend
        lbl_slack = tk.Label(legend_frame, text=" Slack Time ", bg="#95a5a6", fg="white", font=('Arial', 9, 'bold'))
        lbl_slack.pack(side=tk.LEFT, padx=5)
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.gantt_frame.scrollable_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add toolbar for other controls
        toolbar = NavigationToolbar2Tk(canvas, self.gantt_frame.scrollable_frame)
        toolbar.update()
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=10)
        
        # Enable mouse wheel zoom and pan
        self.zoom_factory(ax)
        self.pan_factory(ax)
        
    def on_tab_change(self, event):
        selected_tab = event.widget.select()
        tab_text = event.widget.tab(selected_tab, "text")
        
        if "CPM" in tab_text:
            self.show_cpm_results()
        elif "Network" in tab_text:
            self.show_network_diagram()
        elif "Gantt" in tab_text:
            self.show_gantt_chart()

if __name__ == "__main__":
    root = tk.Tk()
    app = ProjectSchedulingApp(root)
    
    # Bind tab change event
    app.notebook.bind("<<NotebookTabChanged>>", app.on_tab_change)
    
    root.mainloop()
