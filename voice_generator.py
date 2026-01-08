# -*- coding: utf-8 -*-
"""
ElevenLabs Voice Generator Tool
エクセルの台詞リストからElevenLabs TTSで音声ファイルを一括生成するツール
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import threading
import tempfile
import time
from pathlib import Path

import openpyxl
import requests
from pydub import AudioSegment
import pygame

# 設定ファイルのパス
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


class ElevenLabsAPI:
    """ElevenLabs API連携クラス"""
    
    BASE_URL = "https://api.elevenlabs.io/v1"
    
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.headers = {
            "xi-api-key": api_key,
            "Content-Type": "application/json"
        }
    
    def get_voices(self) -> list:
        """利用可能なボイス一覧を取得"""
        try:
            response = requests.get(
                f"{self.BASE_URL}/voices",
                headers=self.headers
            )
            response.raise_for_status()
            data = response.json()
            return data.get("voices", [])
        except Exception as e:
            raise Exception(f"ボイス一覧の取得に失敗しました: {e}")
    
    def generate_speech(self, text: str, voice_id: str) -> bytes:
        """テキストから音声を生成"""
        try:
            response = requests.post(
                f"{self.BASE_URL}/text-to-speech/{voice_id}",
                headers=self.headers,
                json={
                    "text": text,
                    "model_id": "eleven_multilingual_v2",
                    "voice_settings": {
                        "stability": 0.5,
                        "similarity_boost": 0.75
                    }
                }
            )
            response.raise_for_status()
            return response.content
        except Exception as e:
            raise Exception(f"音声生成に失敗しました: {e}")


class ExcelReader:
    """エクセルファイル読み込みクラス"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.workbook = openpyxl.load_workbook(file_path, read_only=False, data_only=True)
        self.sheet = None
        self.cached_data = None
    
    def get_sheet_names(self) -> list:
        """シート名一覧を取得"""
        return self.workbook.sheetnames
    
    def set_sheet(self, sheet_name: str):
        """使用するシートを設定し、データをキャッシュ"""
        self.sheet = self.workbook[sheet_name]
        self.cached_data = []
        for row in self.sheet.iter_rows(values_only=True):
            self.cached_data.append(row)
    
    def get_column_letters(self) -> list:
        """列のアルファベット一覧を取得"""
        if not self.sheet:
            return []
        max_col = self.sheet.max_column
        return [openpyxl.utils.get_column_letter(i) for i in range(1, max_col + 1)]
    
    def _column_index(self, column_letter: str) -> int:
        """列文字を0始まりのインデックスに変換"""
        return openpyxl.utils.column_index_from_string(column_letter) - 1
    
    def get_unique_values_in_column(self, column_letter: str, start_row: int = 2) -> list:
        """指定列のユニークな値を取得"""
        if not self.cached_data:
            return []
        
        col_idx = self._column_index(column_letter)
        values = set()
        
        for row_idx in range(start_row - 1, len(self.cached_data)):
            row = self.cached_data[row_idx]
            if col_idx < len(row) and row[col_idx]:
                values.add(str(row[col_idx]).strip())
        
        return sorted(list(values))
    
    def get_rows_for_character(self, char_column: str, character: str, 
                                dialogue_column: str, filename_column: str, 
                                start_row: int) -> list:
        """特定キャラクターの台詞とファイル名を取得"""
        if not self.cached_data:
            return []
        
        char_idx = self._column_index(char_column)
        dialogue_idx = self._column_index(dialogue_column)
        filename_idx = self._column_index(filename_column)
        
        rows = []
        for row_idx in range(start_row - 1, len(self.cached_data)):
            row = self.cached_data[row_idx]
            
            if char_idx >= len(row):
                continue
                
            char_value = row[char_idx]
            if char_value and str(char_value).strip() == character:
                dialogue = row[dialogue_idx] if dialogue_idx < len(row) else None
                filename = row[filename_idx] if filename_idx < len(row) else None
                if dialogue and filename:
                    rows.append({
                        "dialogue": str(dialogue).strip(),
                        "filename": str(filename).strip()
                    })
        return rows
    
    def close(self):
        self.workbook.close()


class AudioConverter:
    """音声変換クラス"""
    
    @staticmethod
    def mp3_to_wav(mp3_data: bytes, output_path: str):
        """MP3データをWAV (16bit, 44100Hz) に変換して保存"""
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmp:
            tmp.write(mp3_data)
            tmp_path = tmp.name
        
        try:
            audio = AudioSegment.from_mp3(tmp_path)
            audio = audio.set_frame_rate(44100).set_sample_width(2).set_channels(2)
            audio.export(output_path, format="wav")
        finally:
            os.unlink(tmp_path)


class VoiceGeneratorApp:
    """メインアプリケーションクラス"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ElevenLabs Voice Generator")
        self.root.geometry("900x800")
        self.root.resizable(True, True)
        
        # 変数の初期化
        self.api_key = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.char_column = tk.StringVar()
        self.dialogue_column = tk.StringVar()
        self.filename_column = tk.StringVar()
        self.start_row = tk.StringVar(value="2")
        self.output_path = tk.StringVar()
        
        self.excel_reader = None
        self.elevenlabs_api = None
        self.voices = []
        self.characters = []
        self.voice_combos = {}
        
        # pygame初期化（音声再生用）
        pygame.mixer.init()
        
        # 設定を読み込み
        self.load_config()
        
        # UIを構築
        self.build_ui()
    
    def load_config(self):
        """設定ファイルからAPIキーを読み込み"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    self.api_key.set(config.get("api_key", ""))
            except:
                pass
    
    def save_config(self):
        """設定ファイルにAPIキーを保存"""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({"api_key": self.api_key.get()}, f)
        except Exception as e:
            messagebox.showerror("エラー", f"設定の保存に失敗しました: {e}")
    
    def build_ui(self):
        """UIを構築"""
        # メインフレーム（スクロール可能）
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=canvas.yview)
        main_frame = ttk.Frame(canvas, padding="10")
        
        main_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=main_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ===== セクション1: APIキー =====
        section1 = ttk.LabelFrame(main_frame, text="① APIキー設定", padding="10")
        section1.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(section1, text="ElevenLabs APIキー:").pack(anchor=tk.W)
        api_frame = ttk.Frame(section1)
        api_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.api_entry = ttk.Entry(api_frame, textvariable=self.api_key, show="*", width=60)
        self.api_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(api_frame, text="保存", command=self.save_api_key).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(api_frame, text="接続テスト", command=self.test_api_connection).pack(side=tk.LEFT, padx=(5, 0))
        
        # ===== セクション2: エクセルファイル =====
        section2 = ttk.LabelFrame(main_frame, text="② エクセルファイルを選択", padding="10")
        section2.pack(fill=tk.X, pady=(0, 10))
        
        excel_frame = ttk.Frame(section2)
        excel_frame.pack(fill=tk.X)
        
        ttk.Entry(excel_frame, textvariable=self.excel_path, width=60).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(excel_frame, text="参照...", command=self.browse_excel).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(excel_frame, text="読み込み", command=self.load_excel).pack(side=tk.LEFT, padx=(5, 0))
        
        # シート選択
        sheet_frame = ttk.Frame(section2)
        sheet_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(sheet_frame, text="使用するシート:").pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.sheet_name, width=30, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, padx=(5, 10))
        ttk.Button(sheet_frame, text="シートを選択", command=self.select_sheet).pack(side=tk.LEFT)
        
        # ===== セクション3: 列指定 =====
        section3 = ttk.LabelFrame(main_frame, text="③ 列と開始行を指定", padding="10")
        section3.pack(fill=tk.X, pady=(0, 10))
        
        row1 = ttk.Frame(section3)
        row1.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(row1, text="キャラクター名の列:", width=20, anchor=tk.W).pack(side=tk.LEFT)
        self.char_column_combo = ttk.Combobox(row1, textvariable=self.char_column, width=10, state="readonly")
        self.char_column_combo.pack(side=tk.LEFT, padx=(5, 0))
        
        row2 = ttk.Frame(section3)
        row2.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(row2, text="台詞の列:", width=20, anchor=tk.W).pack(side=tk.LEFT)
        self.dialogue_column_combo = ttk.Combobox(row2, textvariable=self.dialogue_column, width=10, state="readonly")
        self.dialogue_column_combo.pack(side=tk.LEFT, padx=(5, 0))
        
        row3 = ttk.Frame(section3)
        row3.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(row3, text="ファイル名の列:", width=20, anchor=tk.W).pack(side=tk.LEFT)
        self.filename_column_combo = ttk.Combobox(row3, textvariable=self.filename_column, width=10, state="readonly")
        self.filename_column_combo.pack(side=tk.LEFT, padx=(5, 0))
        
        row4 = ttk.Frame(section3)
        row4.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(row4, text="データ開始行:", width=20, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(row4, textvariable=self.start_row, width=10).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(row4, text="（ヘッダーが1行目なら2を入力）").pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(section3, text="キャラクター一覧を読み込み", command=self.load_characters).pack(pady=(10, 0))
        
        # ===== セクション4: キャラクター選択 =====
        section4 = ttk.LabelFrame(main_frame, text="④ 書き出すキャラクターを選択（Ctrlキーで複数選択）", padding="10")
        section4.pack(fill=tk.X, pady=(0, 10))
        
        self.char_listbox_frame = ttk.Frame(section4)
        self.char_listbox_frame.pack(fill=tk.X)
        
        self.char_listbox = tk.Listbox(self.char_listbox_frame, selectmode=tk.MULTIPLE, height=6)
        self.char_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        char_scrollbar = ttk.Scrollbar(self.char_listbox_frame, orient=tk.VERTICAL, command=self.char_listbox.yview)
        char_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.char_listbox.config(yscrollcommand=char_scrollbar.set)
        
        # ===== セクション5: ボイス割り当て =====
        section5 = ttk.LabelFrame(main_frame, text="⑤⑥⑦ キャラクターにElevenLabsボイスを割り当て（プレビュー可能）", padding="10")
        section5.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        ttk.Button(section5, text="選択したキャラクターのボイス設定を開始", 
                   command=self.setup_voice_assignment).pack(pady=(0, 10))
        
        voice_canvas = tk.Canvas(section5, height=150)
        voice_scrollbar = ttk.Scrollbar(section5, orient=tk.VERTICAL, command=voice_canvas.yview)
        self.voice_assign_frame = ttk.Frame(voice_canvas)
        
        self.voice_assign_frame.bind(
            "<Configure>",
            lambda e: voice_canvas.configure(scrollregion=voice_canvas.bbox("all"))
        )
        
        voice_canvas.create_window((0, 0), window=self.voice_assign_frame, anchor="nw")
        voice_canvas.configure(yscrollcommand=voice_scrollbar.set)
        
        voice_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        voice_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ===== セクション6: 出力設定 =====
        section6 = ttk.LabelFrame(main_frame, text="⑧ 出力設定", padding="10")
        section6.pack(fill=tk.X, pady=(0, 10))
        
        output_frame = ttk.Frame(section6)
        output_frame.pack(fill=tk.X)
        
        ttk.Label(output_frame, text="出力先フォルダ:").pack(side=tk.LEFT)
        ttk.Entry(output_frame, textvariable=self.output_path, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
        ttk.Button(output_frame, text="参照...", command=self.browse_output).pack(side=tk.LEFT)
        
        self.generate_btn = ttk.Button(section6, text="🎵 音声ファイルを生成", command=self.generate_voices)
        self.generate_btn.pack(pady=(10, 0))
        
        self.progress = ttk.Progressbar(section6, mode="determinate")
        self.progress.pack(fill=tk.X, pady=(10, 0))
        
        self.status_label = ttk.Label(section6, text="")
        self.status_label.pack()
    
    def save_api_key(self):
        """APIキーを保存"""
        self.save_config()
        messagebox.showinfo("完了", "APIキーを保存しました")
    
    def test_api_connection(self):
        """API接続テスト"""
        if not self.api_key.get():
            messagebox.showerror("エラー", "APIキーを入力してください")
            return
        
        try:
            self.elevenlabs_api = ElevenLabsAPI(self.api_key.get())
            self.voices = self.elevenlabs_api.get_voices()
            messagebox.showinfo("成功", f"接続成功！{len(self.voices)}個のボイスが利用可能です")
            self.save_config()
        except Exception as e:
            messagebox.showerror("エラー", str(e))
    
    def browse_excel(self):
        """エクセルファイルを選択"""
        path = filedialog.askopenfilename(
            title="エクセルファイルを選択",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.excel_path.set(path)
    
    def load_excel(self):
        """エクセルファイルを読み込み"""
        if not self.excel_path.get():
            messagebox.showerror("エラー", "エクセルファイルを選択してください")
            return
        
        self.status_label.config(text="エクセルファイルを読み込み中...")
        self.root.update()
        
        try:
            if self.excel_reader:
                self.excel_reader.close()
            
            self.excel_reader = ExcelReader(self.excel_path.get())
            sheet_names = self.excel_reader.get_sheet_names()
            
            self.sheet_combo["values"] = sheet_names
            if sheet_names:
                self.sheet_combo.current(0)
            
            self.char_column_combo["values"] = []
            self.dialogue_column_combo["values"] = []
            self.filename_column_combo["values"] = []
            self.char_column.set("")
            self.dialogue_column.set("")
            self.filename_column.set("")
            
            self.char_listbox.delete(0, tk.END)
            self.characters = []
            
            self.status_label.config(text="")
            messagebox.showinfo("成功", f"エクセルファイルを読み込みました\nシート数: {len(sheet_names)}\n\n使用するシートを選択して「シートを選択」ボタンを押してください")
        except Exception as e:
            self.status_label.config(text="")
            messagebox.showerror("エラー", f"読み込みに失敗しました: {e}")
    
    def select_sheet(self):
        """シートを選択して列情報を読み込み"""
        if not self.excel_reader:
            messagebox.showerror("エラー", "先にエクセルファイルを読み込んでください")
            return
        
        if not self.sheet_name.get():
            messagebox.showerror("エラー", "シートを選択してください")
            return
        
        self.status_label.config(text="シートを読み込み中... しばらくお待ちください")
        self.root.update()
        
        try:
            self.excel_reader.set_sheet(self.sheet_name.get())
            columns = self.excel_reader.get_column_letters()
            
            self.char_column_combo["values"] = columns
            self.dialogue_column_combo["values"] = columns
            self.filename_column_combo["values"] = columns
            
            if columns:
                self.char_column.set(columns[0])
                if len(columns) > 1:
                    self.dialogue_column.set(columns[1])
                if len(columns) > 2:
                    self.filename_column.set(columns[2])
            
            self.char_listbox.delete(0, tk.END)
            self.characters = []
            
            self.status_label.config(text="")
            
            row_count = len(self.excel_reader.cached_data) if self.excel_reader.cached_data else 0
            messagebox.showinfo("成功", f"シート「{self.sheet_name.get()}」を読み込みました\n行数: {row_count}行\n列: {', '.join(columns)}")
        except Exception as e:
            self.status_label.config(text="")
            messagebox.showerror("エラー", f"シートの読み込みに失敗しました: {e}")
    
    def load_characters(self):
        """キャラクター一覧を読み込み"""
        if not self.excel_reader:
            messagebox.showerror("エラー", "先にエクセルファイルを読み込んでください")
            return
        
        if not self.excel_reader.cached_data:
            messagebox.showerror("エラー", "先にシートを選択してください")
            return
        
        if not self.char_column.get():
            messagebox.showerror("エラー", "キャラクター名の列を選択してください")
            return
        
        try:
            start_row = int(self.start_row.get())
        except:
            start_row = 2
        
        self.status_label.config(text="キャラクター一覧を作成中...")
        self.root.update()
        
        try:
            self.characters = self.excel_reader.get_unique_values_in_column(
                self.char_column.get(), start_row
            )
            self.char_listbox.delete(0, tk.END)
            for char in self.characters:
                self.char_listbox.insert(tk.END, char)
            
            self.status_label.config(text="")
            messagebox.showinfo("成功", f"{len(self.characters)}人のキャラクターが見つかりました")
        except Exception as e:
            self.status_label.config(text="")
            messagebox.showerror("エラー", f"読み込みに失敗しました: {e}")
    
    def setup_voice_assignment(self):
        """選択したキャラクターのボイス割り当てUIを構築"""
        selected_indices = self.char_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("エラー", "書き出すキャラクターを選択してください")
            return
        
        if not self.elevenlabs_api:
            try:
                self.elevenlabs_api = ElevenLabsAPI(self.api_key.get())
                self.voices = self.elevenlabs_api.get_voices()
            except Exception as e:
                messagebox.showerror("エラー", f"API接続に失敗しました: {e}")
                return
        
        for widget in self.voice_assign_frame.winfo_children():
            widget.destroy()
        
        selected_chars = [self.characters[i] for i in selected_indices]
        voice_names = [v["name"] for v in self.voices]
        
        self.voice_combos = {}
        
        for i, char in enumerate(selected_chars):
            row_frame = ttk.Frame(self.voice_assign_frame)
            row_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(row_frame, text=f"{char}:", width=20, anchor=tk.W).pack(side=tk.LEFT)
            
            voice_var = tk.StringVar()
            voice_combo = ttk.Combobox(row_frame, textvariable=voice_var, values=voice_names, width=25, state="readonly")
            voice_combo.pack(side=tk.LEFT, padx=(5, 10))
            if voice_names:
                voice_combo.current(0)
            
            self.voice_combos[char] = voice_var
            
            preview_btn = ttk.Button(row_frame, text="▶ プレビュー", 
                                     command=lambda c=char, v=voice_var: self.preview_voice(c, v))
            preview_btn.pack(side=tk.LEFT)
        
        messagebox.showinfo("準備完了", f"{len(selected_chars)}人のキャラクターのボイス設定ができます")
    
    def preview_voice(self, character: str, voice_var: tk.StringVar):
        """選択したボイスでプレビュー再生"""
        if not self.excel_reader:
            messagebox.showerror("エラー", "エクセルファイルを読み込んでください")
            return
        
        voice_name = voice_var.get()
        if not voice_name:
            messagebox.showerror("エラー", "ボイスを選択してください")
            return
        
        voice_id = None
        for v in self.voices:
            if v["name"] == voice_name:
                voice_id = v["voice_id"]
                break
        
        if not voice_id:
            messagebox.showerror("エラー", "ボイスが見つかりません")
            return
        
        try:
            start_row = int(self.start_row.get())
        except:
            start_row = 2
        
        rows = self.excel_reader.get_rows_for_character(
            self.char_column.get(), character,
            self.dialogue_column.get(), self.filename_column.get(),
            start_row
        )
        
        if not rows:
            messagebox.showerror("エラー", f"{character}の台詞が見つかりません")
            return
        
        first_dialogue = rows[0]["dialogue"]
        
        def generate_preview():
            tmp_path = None
            try:
                display_text = first_dialogue[:30] + "..." if len(first_dialogue) > 30 else first_dialogue
                self.status_label.config(text=f"プレビュー生成中: 「{display_text}」")
                mp3_data = self.elevenlabs_api.generate_speech(first_dialogue, voice_id)
                
                with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmp:
                    tmp.write(mp3_data)
                    tmp_path = tmp.name
                
                pygame.mixer.music.load(tmp_path)
                pygame.mixer.music.play()
                
                self.status_label.config(text=f"再生中: {character} - {voice_name}")
                
                while pygame.mixer.music.get_busy():
                    time.sleep(0.1)
                
                self.status_label.config(text="")
                
            except Exception as e:
                self.status_label.config(text="")
            finally:
                if tmp_path:
                    try:
                        time.sleep(0.3)
                        os.unlink(tmp_path)
                    except:
                        pass
        
        threading.Thread(target=generate_preview, daemon=True).start()
    
    def browse_output(self):
        """出力先フォルダを選択"""
        path = filedialog.askdirectory(title="出力先フォルダを選択")
        if path:
            self.output_path.set(path)
    
    def generate_voices(self):
        """音声ファイルを一括生成"""
        if not self.output_path.get():
            messagebox.showerror("エラー", "出力先フォルダを選択してください")
            return
        
        if not self.voice_combos:
            messagebox.showerror("エラー", "キャラクターのボイス設定を行ってください")
            return
        
        if not self.excel_reader:
            messagebox.showerror("エラー", "エクセルファイルを読み込んでください")
            return
        
        os.makedirs(self.output_path.get(), exist_ok=True)
        
        char_voice_map = {}
        for char, voice_var in self.voice_combos.items():
            voice_name = voice_var.get()
            for v in self.voices:
                if v["name"] == voice_name:
                    char_voice_map[char] = v["voice_id"]
                    break
        
        try:
            start_row = int(self.start_row.get())
        except:
            start_row = 2
        
        tasks = []
        for char, voice_id in char_voice_map.items():
            rows = self.excel_reader.get_rows_for_character(
                self.char_column.get(), char,
                self.dialogue_column.get(), self.filename_column.get(),
                start_row
            )
            for row in rows:
                tasks.append({
                    "character": char,
                    "voice_id": voice_id,
                    "dialogue": row["dialogue"],
                    "filename": row["filename"]
                })
        
        if not tasks:
            messagebox.showerror("エラー", "生成する台詞がありません")
            return
        
        if not messagebox.askyesno("確認", f"{len(tasks)}個の音声ファイルを生成しますか？"):
            return
        
        def generate_all():
            self.generate_btn.config(state=tk.DISABLED)
            self.progress["maximum"] = len(tasks)
            self.progress["value"] = 0
            
            success_count = 0
            error_count = 0
            
            for i, task in enumerate(tasks):
                try:
                    self.status_label.config(
                        text=f"生成中 ({i+1}/{len(tasks)}): {task['filename']}"
                    )
                    
                    mp3_data = self.elevenlabs_api.generate_speech(
                        task["dialogue"], task["voice_id"]
                    )
                    
                    filename = task["filename"]
                    if not filename.lower().endswith(".wav"):
                        filename += ".wav"
                    
                    output_file = os.path.join(self.output_path.get(), filename)
                    AudioConverter.mp3_to_wav(mp3_data, output_file)
                    
                    success_count += 1
                    
                except Exception as e:
                    error_count += 1
                    print(f"Error generating {task['filename']}: {e}")
                
                self.progress["value"] = i + 1
                self.root.update()
            
            self.generate_btn.config(state=tk.NORMAL)
            self.status_label.config(text="")
            
            messagebox.showinfo(
                "完了",
                f"音声生成が完了しました\n成功: {success_count}件\nエラー: {error_count}件"
            )
        
        threading.Thread(target=generate_all, daemon=True).start()
    
    def run(self):
        """アプリケーションを実行"""
        self.root.mainloop()
        
        if self.excel_reader:
            self.excel_reader.close()
        pygame.mixer.quit()


if __name__ == "__main__":
    app = VoiceGeneratorApp()
    app.run()
