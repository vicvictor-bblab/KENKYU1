#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# -------------------------------------------------------------------
# CSV Utilities
# -------------------------------------------------------------------

def _load_csv(path: str) -> pd.DataFrame:
    """Read CSV and return cleaned DataFrame for analysis."""
    df = pd.read_csv(path, header=4)

    # Drop unit row if present
    if df.shape[0] > 0 and str(df.iloc[0, 0]).startswith("DataUnit"):
        df = df.drop(index=0).reset_index(drop=True)

    # Remove redundant label column
    if "DataLabel" in df.columns:
        df = df.drop(columns=["DataLabel"])

    # Standardize column names
    rename_map = {
        "Unnamed: 1": "Time",
        "FY[1]": "Force.Fy.1",
        "FZ[2]": "Force.Fz.2",
    }
    df = df.rename(columns=rename_map)

    return df

# ===================================================================
# I. パラメータ定義 (Parameter Definitions)
# ===================================================================
SAMPLING_RATE = 1000.0
FORCE_THRESHOLD_N = 10.0
BASELINE_PERIOD_S = 1.0
FOOT_CONTACT_SD_FACTOR = 5.0

# ===================================================================
# II. GUIアプリケーションクラス (GUI Application Class)
# ===================================================================
class ForceAnalysisApp:
    def __init__(self, root):
        """アプリケーションの初期化とGUIのセットアップ"""
        self.root = root
        self.root.title("高機能・波形分析ツール for LMJ & 投球")
        self.root.geometry("800x750")

        # --- データ管理 ---
        self.results_data = []
        self.current_analysis_result = None
        self.df = None

        # --- スタイル設定 ---
        self._setup_styles()

        # --- 変数定義 ---
        self.analysis_mode = tk.StringVar(value="LMJ")
        self.filepath = tk.StringVar()
        self.subject_name = tk.StringVar()

        # --- GUIの組み立て ---
        self._create_widgets()

    def _setup_styles(self):
        """GUIウィジェットのスタイルを設定"""
        style = ttk.Style()
        style.configure("TLabel", font=("Helvetica", 10))
        style.configure("TButton", font=("Helvetica", 10), padding=5)
        style.configure("TRadiobutton", font=("Helvetica", 10))
        style.configure("TFrame", padding=10)
        style.configure("TLabelframe", padding=10)
        style.configure("TLabelframe.Label", font=("Helvetica", 11, "bold"))
        style.configure("Accent.TButton", foreground="white", background="#0078D7")

    def _create_widgets(self):
        """GUIの全ウィジェットを作成・配置"""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- 上部：コントロールパネル ---
        control_panel = ttk.LabelFrame(main_frame, text="コントロールパネル")
        control_panel.pack(fill=tk.X, pady=(0, 10))
        self._create_control_widgets(control_panel)

        # --- 中部：グラフ表示 ---
        graph_panel = ttk.LabelFrame(main_frame, text="波形グラフ")
        graph_panel.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.fig, self.ax = plt.subplots(figsize=(8, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_panel)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill=tk.BOTH, expand=True)

        # --- 下部：結果とデータ管理 ---
        bottom_panel = ttk.Frame(main_frame)
        bottom_panel.pack(fill=tk.X)
        self._create_result_widgets(bottom_panel)

    def _create_control_widgets(self, parent):
        """コントロールパネル内のウィジェットを作成"""
        # 1. 被験者名
        name_frame = ttk.Frame(parent)
        name_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(name_frame, text="1. 被験者名:", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Entry(name_frame, textvariable=self.subject_name, width=20).pack(side=tk.LEFT)

        # 2. 分析モード
        mode_frame = ttk.Frame(parent)
        mode_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(mode_frame, text="2. 分析モード:", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(mode_frame, text="LMJ", variable=self.analysis_mode, value="LMJ").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="投球", variable=self.analysis_mode, value="投球").pack(side=tk.LEFT, padx=10)

        # 3. ファイル選択
        file_frame = ttk.Frame(parent)
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(file_frame, text="3. ファイル:", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=(0, 10))
        self.file_label = ttk.Label(file_frame, text="選択されていません", width=40, anchor="w")
        self.file_label.pack(side=tk.LEFT)
        ttk.Button(file_frame, text="ファイルを選択...", command=self.select_file).pack(side=tk.LEFT, padx=10)

        # 4. 実行ボタン
        ttk.Button(parent, text="分析実行", command=self.run_analysis, style="Accent.TButton").pack(fill=tk.X, padx=5, pady=10, ipady=5)
    
    def _create_result_widgets(self, parent):
        """結果表示とデータ管理ウィジェットを作成"""
        # 結果表示
        result_frame = ttk.LabelFrame(parent, text="分析結果")
        result_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,10))
        self.result_text = tk.Text(result_frame, height=4, width=40, state="disabled", font=("Courier", 10))
        self.result_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # データ管理
        manage_frame = ttk.LabelFrame(parent, text="データ管理")
        manage_frame.pack(side=tk.RIGHT, fill=tk.Y)
        ttk.Button(manage_frame, text="この結果をリストに追加", command=self.add_result_to_list).pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(manage_frame, text="全データをExcelに出力", command=self.export_to_excel).pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(manage_frame, text="終了", command=self.confirm_exit).pack(fill=tk.X, padx=5, pady=2)
        self.saved_count_label = ttk.Label(manage_frame, text="保存済み: 0件")
        self.saved_count_label.pack(pady=5)

    def select_file(self):
        """ファイル選択ダイアログを開き、ファイルパスを取得・表示する"""
        filepath = filedialog.askopenfilename(
            title="CSVファイルを選択", filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        if filepath:
            self.filepath.set(filepath)
            self.file_label.config(text=f"{os.path.basename(filepath)}")

    def run_analysis(self):
        """分析のメインロジックを実行"""
        # 入力チェック
        if not self.subject_name.get():
            messagebox.showerror("入力エラー", "被験者名を入力してください。")
            return
        if not self.filepath.get():
            messagebox.showerror("入力エラー", "分析するCSVファイルを選択してください。")
            return
        
        # データ読み込み
        try:
            # Excelで5行目がヘッダー(DataLabel)
            self.df = _load_csv(self.filepath.get())
        except Exception as e:
            messagebox.showerror("ファイル読込エラー", f"ファイルの読み込みに失敗しました。\nエラー: {e}")
            return
            
        # 分析モードに応じて処理を分岐
        mode = self.analysis_mode.get()
        if mode == "LMJ":
            self.analyze_lmj()
        elif mode == "投球":
            self.analyze_throwing()

    def analyze_lmj(self):
        """LMJの分析ロジック"""
        force_col = 'Force.Fy.1'
        time_col = 'Time'
        if force_col not in self.df.columns or time_col not in self.df.columns:
            messagebox.showerror("エラー", f"必要な列 '{force_col}' または '{time_col}' が見つかりません。")
            return

        force = self.df[force_col].abs()
        candidates = self.df.index[force > FORCE_THRESHOLD_N].tolist()
        
        if not candidates:
            messagebox.showinfo("情報", "分析区間が見つかりませんでした (閾値を超えませんでした)。")
            return
            
        # 閾値を超えている区間の開始と終了を特定
        start_idx = candidates[0]
        end_idx = candidates[-1]

        self.calculate_and_display(start_idx, end_idx, force_col, time_col)

    def analyze_throwing(self):
        """投球の分析ロジック"""
        axis_foot_col = 'Force.Fy.1'
        lead_foot_col = 'Force.Fz.2'
        time_col = 'Time'

        # 必要な列の存在チェック
        for col in [axis_foot_col, lead_foot_col, time_col]:
            if col not in self.df.columns:
                messagebox.showerror("エラー", f"必要な列 '{col}' が見つかりません。")
                return

        # --- 開始点の特定 ---
        axis_force_abs = self.df[axis_foot_col].abs()
        start_candidates_idx = np.where(axis_force_abs > FORCE_THRESHOLD_N)[0]

        if len(start_candidates_idx) == 0:
            messagebox.showinfo("情報", "開始点が見つかりませんでした (閾値を超えませんでした)。")
            return
        
        start_idx = self.choose_index_dialog(start_candidates_idx, "開始点")
        if start_idx is None: return # ユーザーがキャンセルした場合

        # --- 終了点（フットコンタクト）の特定 ---
        lead_force = self.df[lead_foot_col]
        baseline_end_frame = int(BASELINE_PERIOD_S * SAMPLING_RATE)
        baseline_data = lead_force.iloc[:baseline_end_frame]
        baseline_mean = baseline_data.mean()
        baseline_sd = baseline_data.std()
        contact_threshold = baseline_mean + FOOT_CONTACT_SD_FACTOR * baseline_sd
        
        # 開始点以降で閾値を超えた最初の点
        end_candidates_idx = np.where(lead_force.iloc[start_idx:] > contact_threshold)[0]
        if len(end_candidates_idx) == 0:
            messagebox.showinfo("情報", "終了点（フットコンタクト）が見つかりませんでした。")
            return

        # インデックスを元のDataFrameのインデックスに変換
        end_idx = end_candidates_idx[0] + start_idx
        
        self.calculate_and_display(start_idx, end_idx, axis_foot_col, time_col)

    def calculate_and_display(self, start_idx, end_idx, force_col, time_col):
        """指定された区間で計算、結果表示、グラフ描画を行う"""
        analysis_df = self.df.iloc[start_idx:end_idx+1]
        
        # 計算
        peak_force = analysis_df[force_col].abs().max()
        # 力積の計算 (台形則)
        impulse = np.trapz(y=analysis_df[force_col].abs(), x=analysis_df[time_col])

        # 結果を辞書に格納
        self.current_analysis_result = {
            "被験者名": self.subject_name.get(),
            "分析モード": self.analysis_mode.get(),
            "ファイル名": os.path.basename(self.filepath.get()),
            "ピークフォース(N)": round(peak_force, 2),
            "力積(N・s)": round(impulse, 2),
            "開始時間(s)": self.df.loc[start_idx, time_col],
            "終了時間(s)": self.df.loc[end_idx, time_col],
        }

        # 結果をテキストボックスに表示
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", tk.END)
        result_str = (
            f"ピークフォース: {self.current_analysis_result['ピークフォース(N)']} N\n"
            f"力積        : {self.current_analysis_result['力積(N・s)']} Ns\n"
            f"分析区間    : {self.current_analysis_result['開始時間(s)']}s - {self.current_analysis_result['終了時間(s)']}s"
        )
        self.result_text.insert(tk.END, result_str)
        self.result_text.config(state="disabled")

        # グラフを描画
        self.plot_waveform(start_idx, end_idx, force_col, time_col)

    def plot_waveform(self, start_idx, end_idx, force_col, time_col):
        """波形と分析区間をグラフに描画"""
        self.ax.clear()
        
        time_data = self.df[time_col]
        force_data = self.df[force_col]
        start_time = self.df.loc[start_idx, time_col]
        end_time = self.df.loc[end_idx, time_col]

        self.ax.plot(time_data, force_data, label=force_col, color='gray')
        
        # 分析区間の波形をハイライト
        self.ax.plot(time_data[start_idx:end_idx+1], force_data[start_idx:end_idx+1], color='dodgerblue', linewidth=2)
        
        # 開始・終了の垂直線
        self.ax.axvline(x=start_time, color='green', linestyle='--', label=f'開始: {start_time:.3f}s')
        self.ax.axvline(x=end_time, color='red', linestyle='--', label=f'終了: {end_time:.3f}s')
        
        # 分析区間を塗りつぶし
        self.ax.fill_between(time_data, self.ax.get_ylim()[0], self.ax.get_ylim()[1], 
                             where=(time_data >= start_time) & (time_data <= end_time), 
                             color='dodgerblue', alpha=0.1)

        self.ax.set_title(f"波形グラフ ({self.analysis_mode.get()})")
        self.ax.set_xlabel("時間 (s)")
        self.ax.set_ylabel("力 (N)")
        self.ax.legend()
        self.ax.grid(True, linestyle=':')
        self.fig.tight_layout()
        self.canvas.draw()

    def choose_index_dialog(self, indices, event_name):
        """複数の候補からユーザーに1つを選択させるダイアログを表示"""
        if len(indices) == 1:
            return indices[0]

        dialog = tk.Toplevel(self.root)
        dialog.title(f"{event_name}の候補を選択")
        dialog.geometry("350x150")
        
        ttk.Label(dialog, text=f"{event_name}の候補が複数見つかりました。").pack(pady=5)
        ttk.Label(dialog, text="使用する時間（秒）を選択してください。").pack(pady=5)
        
        times = [f"{self.df.loc[i, 'Time']:.4f}" for i in indices]
        
        selected_time = tk.StringVar()
        combobox = ttk.Combobox(dialog, textvariable=selected_time, values=times, state="readonly")
        combobox.pack(pady=5)
        combobox.set(times[0])
        
        result_index = tk.IntVar()

        def on_ok():
            chosen_time_str = selected_time.get()
            chosen_time_float = float(chosen_time_str)
            # 選択された時間から最も近いインデックスを探す
            result_index.set(self.df.index[abs(self.df['Time'] - chosen_time_float) < 1e-9][0])
            dialog.destroy()
            
        def on_cancel():
            result_index.set(-1) # キャンセルを示す
            dialog.destroy()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="キャンセル", command=on_cancel).pack(side=tk.LEFT)
        
        self.root.wait_window(dialog) # ダイアログが閉じるのを待つ
        
        final_index = result_index.get()
        return final_index if final_index != -1 else None

    def add_result_to_list(self):
        """現在表示されている分析結果をリストに蓄積"""
        if self.current_analysis_result:
            self.results_data.append(self.current_analysis_result)
            count = len(self.results_data)
            self.saved_count_label.config(text=f"保存済み: {count}件")
            messagebox.showinfo("成功", "結果をリストに追加しました。")
            self.current_analysis_result = None # 重複追加を防ぐ
        else:
            messagebox.showwarning("注意", "リストに追加する分析結果がありません。\n先に分析を実行してください。")
    
    def export_to_excel(self):
        """蓄積した全データをExcelファイルに出力"""
        if not self.results_data:
            messagebox.showwarning("注意", "Excelに出力するデータがありません。")
            return
            
        filepath = filedialog.asksaveasfilename(
            title="Excelファイルとして保存",
            defaultextension=".xlsx",
            filetypes=[("Excelファイル", "*.xlsx")]
        )
        if filepath:
            try:
                df_to_export = pd.DataFrame(self.results_data)
                df_to_export.to_excel(filepath, index=False)
                messagebox.showinfo("成功", f"全{len(self.results_data)}件のデータを\n{filepath}\nに保存しました。")
            except Exception as e:
                messagebox.showerror("保存エラー", f"Excelファイルの保存に失敗しました。\nエラー: {e}")

    def confirm_exit(self):
        """アプリ終了前に確認ダイアログを表示"""
        ok = messagebox.askyesno(
            "終了確認",
            "本当に終了していいですか？\nデータの保存は完了していますか？"
        )
        if ok:
            self.root.destroy()

# ===================================================================
# III. アプリケーションの実行 (Application Execution)
# ===================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = ForceAnalysisApp(root)
    root.mainloop()
