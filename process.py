import sys
import os
import json
import csv
import requests
import traceback
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QHeaderView, QGroupBox, QComboBox, QMessageBox, 
                             QProgressBar, QTextEdit, QTreeWidget, QTreeWidgetItem,
                             QFrame, QFileDialog)
from PyQt6.QtCore import Qt, pyqtSignal, QThread
from PyQt6.QtGui import QFont, QColor

# --- [1] AI 엔진 ---
class AIWorker(QThread):
    finished = pyqtSignal(list, str) 
    error = pyqtSignal(str)

    def __init__(self, hierarchy_data, model_name, process_info):
        super().__init__()
        self.hierarchy_data = hierarchy_data
        self.model_name = model_name
        self.process_info = process_info

    def run(self):
        url = "http://localhost:11434/api/chat"
        prompt = f"""당신은 보안이 철저한 폐쇄망(Offline) 환경의 최고급 프로세스 혁신 컨설턴트입니다. 
        계층형 AS-IS 프로세스를 분석하여 최적의 TO-BE 혁신안을 제안하세요.

        [프로젝트 정보] {self.process_info}
        [입력 데이터]
        {self.hierarchy_data}
        
        [혁신 지시사항 - 절대 준수]
        1. 계층 유지: 입력된 No와 계층 구조를 1:1로 매칭. (상위 프로세스는 수치/유형 비움 '-')
        2. 전략적 유형 전환: 사람 판단이 필요하면 '전산(ERP/NF/기타)', AI/봇이 100% 대체하면 '자동화'.
        3. 폐쇄망 기술 제안: n8n 로컬, Apache Airflow, EasyOCR, 로컬 LLM(Ollama) 등 명시.
        4. 산술 및 인원 분리 (매우 중요): 
           - 공식: [총공수] = [횟수] x [건당시간] x [투입인원]
           - '자동화' 적용 시 사람의 개입이 사라지므로 [투입인원]을 0으로 만드세요. 대신 [건당시간]에 봇(Bot)의 처리 속도를 적어 시스템 공수를 산출하세요.
           - '전산화' 적용 시 처리 속도가 빨라지므로 [건당시간]이나 [투입인원]을 줄이세요.
        
        [출력 형식]
        리포트: (폐쇄망 전산화/자동화 도입 전략 및 인력 재배치 효과를 3~4줄로 상세 요약)
        데이터:
        No|단위 프로세스명|횟수|건당시간|투입인원|총공수|유형|업무방법"""

        try:
            res = requests.post(url, json={"model": self.model_name, "messages": [{"role": "user", "content": prompt}], "stream": False}, timeout=400)
            if res.status_code == 200:
                content = res.json()['message']['content']
                report_part = content.split("데이터:")[0].replace("리포트:", "").strip()
                data_part = content.split("데이터:")[1].strip()
                
                rows = []
                for line in data_part.split("\n"):
                    line = line.strip()
                    if not line or "---" in line or "단위 프로세스명" in line: continue
                    cols = [c.strip() for c in line.strip("|").split("|")]
                    if len(cols) >= 8: rows.append(cols[:8])
                self.finished.emit(rows, report_part)
            else:
                self.error.emit(f"HTTP {res.status_code} 오류 발생")
        except Exception as e:
            self.error.emit(f"통신 에러: {str(e)}")

# --- [2] 메인 UI 대시보드 ---
class ProcessAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.secu_path = "C:/secu"
        self.init_ui()
        self.apply_styles()

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f4f6f8; }
            QGroupBox { font-weight: bold; font-size: 13px; border: 2px solid #adb5bd; border-radius: 8px; margin-top: 12px; padding: 12px; background-color: white; }
            QTreeWidget::item { min-height: 35px; padding: 2px; }
            QTreeWidget { border: 1px solid #dee2e6; alternate-background-color: #f8f9fa; font-size: 12px; }
            QHeaderView::section { background-color: #e9ecef; font-weight: bold; border: 1px solid #dee2e6; padding: 4px; }
            QPushButton { border-radius: 6px; padding: 10px; font-weight: bold; font-size: 13px; }
            QComboBox { min-height: 25px; padding: 2px; }
            #btn_run { background-color: #3b5bdb; color: white; }
            #btn_run:hover { background-color: #364fc7; }
            #btn_run:disabled { background-color: #748ffc; color: #e9ecef; }
            #btn_save { background-color: #2f9e44; color: white; }
            #btn_load { background-color: #0b7285; color: white; }
            #btn_csv { background-color: #f59f00; color: white; }
            #dashboard_frame { background-color: #212529; border-radius: 10px; padding: 15px; }
            #dash_title { color: white; font-size: 15px; font-weight: bold; margin-bottom: 5px; }
            #dash_text { color: #f8f9fa; font-weight: bold; font-size: 13px; margin-bottom: 2px; }
            #dash_human { color: #38d9a9; font-weight: bold; font-size: 14px; margin-left: 10px; }
            #dash_ai { color: #fcc419; font-weight: bold; font-size: 14px; margin-left: 10px; margin-bottom: 5px; }
            #dash_highlight { color: #ff8787; font-weight: bold; font-size: 15px; margin-bottom: 2px; margin-top: 5px; }
            #dash_money { color: #69db7c; font-weight: bold; font-size: 16px; margin-bottom: 2px; }
            #dash_time { color: #4dabf7; font-weight: bold; font-size: 16px; margin-bottom: 2px; }
            #dash_digi { color: #74c0fc; font-weight: bold; font-size: 14px; margin-bottom: 2px; }
            #dash_auto { color: #b197fc; font-weight: bold; font-size: 14px; }
        """)

    def init_ui(self):
        self.setWindowTitle("🚀 프로세스 혁신 분석기 V44 (전환율 KPI 완전 복구판)")
        self.setGeometry(20, 20, 1750, 1000)
        
        container = QWidget()
        self.setCentralWidget(container)
        main_layout = QVBoxLayout(container)

        meta_group = QGroupBox("📋 프로젝트 메타데이터 & 환경 설정")
        meta_layout = QHBoxLayout()
        self.proj_name = QLineEdit("전사 표준 프로세스 지능화 사업"); self.dept = QLineEdit("DX추진팀")
        self.author = QLineEdit("관리자"); self.wage_input = QLineEdit("")
        self.wage_input.setPlaceholderText("예: 45000 (입력 시 금액 변환)")
        self.wage_input.setFixedWidth(180) 
        self.model_combo = QComboBox(); self.model_combo.addItems(["gemma3:12b", "gemma3:4b"])
        
        for t, w in [("프로세스명:", self.proj_name), ("부서:", self.dept), ("작성자:", self.author), 
                     ("💰 평균 시급/인건비(원):", self.wage_input), ("모델:", self.model_combo)]:
            meta_layout.addWidget(QLabel(t)); meta_layout.addWidget(w)
        
        self.wage_input.textChanged.connect(self.calc_totals) 
        meta_group.setLayout(meta_layout); main_layout.addWidget(meta_group)

        trees_layout = QHBoxLayout()
        asis_box = QGroupBox("🔴 AS-IS Hierarchy (현재 수기/전산 프로세스)"); asis_v = QVBoxLayout()
        self.asis_tree = self.create_tree(); asis_v.addWidget(self.asis_tree)
        
        tree_btn_layout = QHBoxLayout()
        b_main = QPushButton("➕ 상위 프로세스 추가"); b_main.clicked.connect(self.add_main_process)
        b_sub = QPushButton("↳ 하위 세부 추가"); b_sub.clicked.connect(self.add_sub_process)
        b_del = QPushButton("➖ 선택 삭제"); b_del.clicked.connect(self.delete_item)
        tree_btn_layout.addWidget(b_main); tree_btn_layout.addWidget(b_sub); tree_btn_layout.addWidget(b_del)
        asis_v.addLayout(tree_btn_layout); asis_box.setLayout(asis_v)
        
        tobe_box = QGroupBox("🟢 TO-BE Innovation (AI 전략 설계안)"); tobe_v = QVBoxLayout()
        self.tobe_tree = self.create_tree(); tobe_v.addWidget(self.tobe_tree)
        tobe_box.setLayout(tobe_v)
        
        self.asis_tree.verticalScrollBar().valueChanged.connect(self.tobe_tree.verticalScrollBar().setValue)
        self.tobe_tree.verticalScrollBar().valueChanged.connect(self.asis_tree.verticalScrollBar().setValue)

        trees_layout.addWidget(asis_box, 1); trees_layout.addWidget(tobe_box, 1)
        main_layout.addLayout(trees_layout, stretch=3)

        analysis_layout = QHBoxLayout()
        report_group = QGroupBox("📝 AI 혁신 전략 분석 보고서")
        report_v = QVBoxLayout()
        self.analysis_text = QTextEdit(); self.analysis_text.setReadOnly(True)
        self.analysis_text.setFixedHeight(230); self.analysis_text.setFont(QFont("Malgun Gothic", 11))
        report_v.addWidget(self.analysis_text); report_group.setLayout(report_v)
        
        dash_frame = QFrame(); dash_frame.setObjectName("dashboard_frame"); dash_v = QVBoxLayout()
        
        self.lbl_asis_tot = QLabel("🔴 AS-IS 인력 총공수: 0.0 hr"); self.lbl_asis_tot.setObjectName("dash_text")
        label_tobe_header = QLabel("▼ TO-BE 업무 이관 배분 ▼")
        label_tobe_header.setStyleSheet("color: #ced4da; font-size: 12px; margin-top: 5px; margin-bottom: 2px;")
        self.lbl_tobe_human = QLabel("🧑‍💻 잔여 인력 공수: 0.0 hr"); self.lbl_tobe_human.setObjectName("dash_human")
        self.lbl_tobe_sys = QLabel("🧠 AI 시스템 공수: 0.0 hr"); self.lbl_tobe_sys.setObjectName("dash_ai")
        self.lbl_saving = QLabel("🔥 총 인력 절감: 0.0 hr (0.0%)"); self.lbl_saving.setObjectName("dash_highlight")
        
        separator = QLabel("───────────────────"); separator.setStyleSheet("color: gray;")
        
        # 🔥 V44: 누락되었던 KPI 라벨 복구
        self.lbl_digi_rate = QLabel("💻 전산화 전환율: 0.0%"); self.lbl_digi_rate.setObjectName("dash_digi")
        self.lbl_auto_rate = QLabel("🤖 AI 자동화율: 0.0%"); self.lbl_auto_rate.setObjectName("dash_auto")
        
        self.lbl_roi_month = QLabel("⏱️ 월간 절감 공수: 0.0 hr"); self.lbl_roi_month.setObjectName("dash_time")
        self.lbl_roi_year = QLabel("🚀 연간 절감 공수: 0.0 hr"); self.lbl_roi_year.setObjectName("dash_time")
        
        dash_title = QLabel("[재무 및 혁신 성과 KPI]"); dash_title.setObjectName("dash_title")
        dash_v.addWidget(dash_title); dash_v.addWidget(self.lbl_asis_tot)
        dash_v.addWidget(label_tobe_header); dash_v.addWidget(self.lbl_tobe_human); dash_v.addWidget(self.lbl_tobe_sys)
        dash_v.addWidget(self.lbl_saving)
        dash_v.addWidget(separator)
        
        # 🔥 V44: 대시보드 레이아웃에 KPI 추가
        dash_v.addWidget(self.lbl_digi_rate)
        dash_v.addWidget(self.lbl_auto_rate)
        dash_v.addWidget(self.lbl_roi_month); dash_v.addWidget(self.lbl_roi_year)
        dash_v.addStretch(); dash_frame.setLayout(dash_v)

        analysis_layout.addWidget(report_group, stretch=4)
        analysis_layout.addWidget(dash_frame, stretch=1)
        main_layout.addLayout(analysis_layout, stretch=1)

        bottom_v = QVBoxLayout()
        self.progress = QProgressBar(); self.progress.setVisible(False); self.progress.setFixedHeight(15)
        bottom_v.addWidget(self.progress)
        
        bottom_btns = QHBoxLayout()
        self.btn_run = QPushButton("✨ Gemma 3 12B 정밀 계층 분석 실행"); self.btn_run.setObjectName("btn_run")
        self.btn_run.clicked.connect(self.run_ai); self.btn_run.setFixedHeight(45)
        self.btn_csv = QPushButton("📊 엑셀(CSV) 추출"); self.btn_csv.setObjectName("btn_csv")
        self.btn_csv.clicked.connect(self.export_csv); self.btn_csv.setFixedHeight(45)
        self.btn_load = QPushButton("📂 JSON 불러오기"); self.btn_load.setObjectName("btn_load")
        self.btn_load.clicked.connect(self.load_data); self.btn_load.setFixedHeight(45)
        self.btn_save = QPushButton("💾 전체 상태 JSON 저장"); self.btn_save.setObjectName("btn_save")
        self.btn_save.clicked.connect(self.save_data); self.btn_save.setFixedHeight(45)
        
        bottom_btns.addStretch()
        bottom_btns.addWidget(self.btn_run); bottom_btns.addWidget(self.btn_csv)
        bottom_btns.addWidget(self.btn_load); bottom_btns.addWidget(self.btn_save)
        
        bottom_v.addLayout(bottom_btns); main_layout.addLayout(bottom_v)

        self.asis_tree.itemChanged.connect(self.on_item_changed)
        self.tobe_tree.itemChanged.connect(self.on_item_changed)
        self.load_default_samples()

    def create_tree(self):
        tree = QTreeWidget()
        tree.setColumnCount(8)
        tree.setHeaderLabels(["No", "단위 프로세스명", "횟수(월)", "건당시간(hr)", "인원(명)", "총공수(hr)", "유형", "업무방법"])
        tree.setColumnWidth(0, 50); tree.setColumnWidth(1, 250); tree.setColumnWidth(2, 60)
        tree.setColumnWidth(3, 70); tree.setColumnWidth(4, 60); tree.setColumnWidth(5, 75); tree.setColumnWidth(6, 90)
        tree.header().setStretchLastSection(True); tree.setAlternatingRowColors(True)
        return tree

    def create_combo(self, default_text="수기", is_tobe=False):
        cb = QComboBox()
        cb.addItems(["수기", "전산(ERP)", "전산(NF)", "전산(기타)", "자동화"])
        cb.setCurrentText(default_text)
        
        if is_tobe:
            if "자동화" in default_text: cb.setStyleSheet("color: #9775fa; font-weight: bold;")
            elif "전산" in default_text: cb.setStyleSheet("color: #339af0; font-weight: bold;")
            
            def update_color(text, box=cb):
                if "자동화" in text: box.setStyleSheet("color: #9775fa; font-weight: bold;")
                elif "전산" in text: box.setStyleSheet("color: #339af0; font-weight: bold;")
                else: box.setStyleSheet("color: black; font-weight: normal;")
            cb.currentTextChanged.connect(update_color)
            
        cb.currentIndexChanged.connect(self.calc_totals)
        return cb

    def on_item_changed(self, item, column):
        if column in [2, 3, 4]:
            tree = item.treeWidget()
            try:
                freq = float(item.text(2)) if item.text(2).replace('.','',1).isdigit() else 0.0
                time_per = float(item.text(3)) if item.text(3).replace('.','',1).isdigit() else 0.0
                persons = float(item.text(4)) if item.text(4).replace('.','',1).isdigit() else 0.0
                total = freq * time_per * persons
                tree.blockSignals(True)
                item.setText(5, f"{total:.1f}")
                tree.blockSignals(False)
            except ValueError: pass
        self.calc_totals()

    def get_all_tree_items(self, tree):
        items = []
        def recursive_extract(parent_item):
            for i in range(parent_item.childCount()):
                child = parent_item.child(i)
                items.append(child)
                recursive_extract(child)
        for i in range(tree.topLevelItemCount()):
            top_item = tree.topLevelItem(i)
            items.append(top_item)
            recursive_extract(top_item)
        return items

    # 🔥 V44: 전산화율, 자동화율 계산 로직 완벽 복구
    def calc_totals(self):
        a_s = 0.0; t_s_human = 0.0; t_s_sys = 0.0      
        auto_base_hr = 0.0; digi_base_hr = 0.0 
        
        tobe_items = self.get_all_tree_items(self.tobe_tree)
        tobe_dict = {item.text(0): item for item in tobe_items}
        asis_items = self.get_all_tree_items(self.asis_tree)
        
        for item_a in asis_items:
            if item_a.childCount() == 0: 
                no_val = item_a.text(0)
                val_a = item_a.text(5)
                hr_a = float(val_a) if val_a and val_a.replace('.','',1).isdigit() else 0.0
                a_s += hr_a

                w_type_a = self.asis_tree.itemWidget(item_a, 6)
                type_a = w_type_a.currentText() if w_type_a else item_a.text(6)

                item_t = tobe_dict.get(no_val)
                if item_t:
                    val_t = item_t.text(5)
                    hr_t = float(val_t) if val_t and val_t.replace('.','',1).isdigit() else 0.0
                    
                    w_type_t = self.tobe_tree.itemWidget(item_t, 6)
                    type_t = w_type_t.currentText() if w_type_t else item_t.text(6)

                    # TO-BE가 자동화/전산화 되었을 때, AS-IS의 기존 공수를 가중치로 전환율 합산
                    if "자동화" in type_t: 
                        auto_base_hr += hr_a
                        t_s_sys += hr_t 
                    else: 
                        t_s_human += hr_t 
                        if "수기" in type_a and "전산" in type_t: 
                            digi_base_hr += hr_a

        diff = a_s - t_s_human
        rate = (diff / a_s * 100) if a_s > 0 else 0
        auto_rate = (auto_base_hr / a_s * 100) if a_s > 0 else 0
        digi_rate = (digi_base_hr / a_s * 100) if a_s > 0 else 0
        
        self.lbl_asis_tot.setText(f"🔴 AS-IS 인력 총공수: {a_s:.1f} hr")
        self.lbl_tobe_human.setText(f"🧑‍💻 잔여 인력 공수: {t_s_human:.1f} hr")
        self.lbl_tobe_sys.setText(f"🧠 AI 시스템 공수: {t_s_sys:.1f} hr")
        self.lbl_saving.setText(f"🔥 총 인력 절감: {diff:.1f} hr ({rate:.1f}%)")
        
        # 복구된 라벨 업데이트
        self.lbl_digi_rate.setText(f"💻 전산화 전환율: {digi_rate:.1f}%")
        self.lbl_auto_rate.setText(f"🤖 AI 자동화율: {auto_rate:.1f}%")

        try: 
            wage_text = self.wage_input.text().replace(',', '').strip()
            wage = float(wage_text) if wage_text else 0.0
        except: wage = 0.0
        
        if wage > 0:
            monthly_saved = diff * wage; yearly_saved = monthly_saved * 12
            self.lbl_roi_month.setText(f"💰 월간 절감액: ₩ {int(monthly_saved):,}")
            self.lbl_roi_year.setText(f"🚀 연간 기대효과: ₩ {int(yearly_saved):,}")
            self.lbl_roi_month.setObjectName("dash_money")
            self.lbl_roi_year.setObjectName("dash_money")
        else:
            self.lbl_roi_month.setText(f"⏱️ 월간 절감 공수: {diff:.1f} hr")
            self.lbl_roi_year.setText(f"🚀 연간 절감 공수: {diff * 12:.1f} hr")
            self.lbl_roi_month.setObjectName("dash_time")
            self.lbl_roi_year.setObjectName("dash_time")
            
        self.lbl_roi_month.setStyleSheet(self.styleSheet())
        self.lbl_roi_year.setStyleSheet(self.styleSheet())

    def add_main_process(self):
        count = self.asis_tree.topLevelItemCount() + 1
        item = QTreeWidgetItem(self.asis_tree, [str(count), "상위 프로세스 입력", "-", "-", "-", "-", "-", "하위 참조"])
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
        self.asis_tree.expandItem(item)

    def add_sub_process(self):
        parent = self.asis_tree.currentItem()
        if not parent: return
        p_no = parent.text(0); sub_count = parent.childCount() + 1
        child = QTreeWidgetItem(parent, [f"{p_no}-{sub_count}", "세부 업무 입력", "1", "1.0", "1", "1.0", "수기", ""])
        child.setFlags(child.flags() | Qt.ItemFlag.ItemIsEditable)
        self.asis_tree.setItemWidget(child, 6, self.create_combo("수기"))
        self.asis_tree.expandItem(parent); self.calc_totals()

    def delete_item(self):
        root = self.asis_tree.invisibleRootItem()
        for item in self.asis_tree.selectedItems(): (item.parent() or root).removeChild(item)
        self.calc_totals()

    def load_default_samples(self):
        self.asis_tree.blockSignals(True)
        
        # 1. 자재/물류
        m1 = QTreeWidgetItem(self.asis_tree, ["1", "자재 입고 및 수입 검사 처리", "-", "-", "-", "-", "-", "하위 단계 참조"])
        m1.setFlags(m1.flags() | Qt.ItemFlag.ItemIsEditable)
        s1_1 = QTreeWidgetItem(m1, ["1-1", "납품명세서/송장 종이 수령 및 분류", "200", "0.1", "2", "40.0", "수기", "2명의 담당자가 매일 수동 파일링"])
        s1_1.setFlags(s1_1.flags() | Qt.ItemFlag.ItemIsEditable); self.asis_tree.setItemWidget(s1_1, 6, self.create_combo("수기"))
        s1_2 = QTreeWidgetItem(m1, ["1-2", "거래명세서 데이터 ERP 수동 입력", "200", "0.25", "3", "150.0", "전산(ERP)", "3명이 종이 명세서를 보며 건별 타이핑"])
        s1_2.setFlags(s1_2.flags() | Qt.ItemFlag.ItemIsEditable); self.asis_tree.setItemWidget(s1_2, 6, self.create_combo("전산(ERP)"))
        
        # 2. 생산/설비
        m2 = QTreeWidgetItem(self.asis_tree, ["2", "생산 실적 및 설비 가동 현황 관리", "-", "-", "-", "-", "-", "하위 단계 참조"])
        m2.setFlags(m2.flags() | Qt.ItemFlag.ItemIsEditable)
        s2_1 = QTreeWidgetItem(m2, ["2-1", "현장 작업지시서 NF 종이 출력 및 배포", "30", "0.5", "1", "15.0", "전산(NF)", "NF 시스템 접속하여 PDF 출력 후 현장 배포"])
        s2_1.setFlags(s2_1.flags() | Qt.ItemFlag.ItemIsEditable); self.asis_tree.setItemWidget(s2_1, 6, self.create_combo("전산(NF)"))
        s2_2 = QTreeWidgetItem(m2, ["2-2", "설비 알람 및 에러 코드 수기 메모", "100", "0.2", "2", "40.0", "수기", "알람 발생 시 현장 패널 확인 후 수첩에 기록"])
        s2_2.setFlags(s2_2.flags() | Qt.ItemFlag.ItemIsEditable); self.asis_tree.setItemWidget(s2_2, 6, self.create_combo("수기"))

        # 3. 재무/회계
        m3 = QTreeWidgetItem(self.asis_tree, ["3", "월간 재무/회계 마감 및 비용 정산", "-", "-", "-", "-", "-", "하위 단계 참조"])
        m3.setFlags(m3.flags() | Qt.ItemFlag.ItemIsEditable)
        s3_1 = QTreeWidgetItem(m3, ["3-1", "법인카드 종이 영수증 취합 및 지출결의 대조", "150", "0.2", "2", "60.0", "수기", "부서별 영수증 실물과 엑셀 내역 수동 대조"])
        s3_1.setFlags(s3_1.flags() | Qt.ItemFlag.ItemIsEditable); self.asis_tree.setItemWidget(s3_1, 6, self.create_combo("수기"))
        s3_2 = QTreeWidgetItem(m3, ["3-2", "매출 세금계산서 국세청 발행 및 이메일 발송", "100", "0.15", "1", "15.0", "전산(기타)", "홈택스 로그인 후 건별 발행 및 PDF 다운 후 메일 발송"])
        s3_2.setFlags(s3_2.flags() | Qt.ItemFlag.ItemIsEditable); self.asis_tree.setItemWidget(s3_2, 6, self.create_combo("전산(기타)"))

        # 4. 인사/총무
        m4 = QTreeWidgetItem(self.asis_tree, ["4", "임직원 근태 관리 및 급여 명세 처리", "-", "-", "-", "-", "-", "하위 단계 참조"])
        m4.setFlags(m4.flags() | Qt.ItemFlag.ItemIsEditable)
        s4_1 = QTreeWidgetItem(m4, ["4-1", "부서별 연차/초과근무 수기 대장 취합", "50", "0.3", "1", "15.0", "수기", "그룹웨어 데이터와 수기 대장 크로스 체크"])
        s4_1.setFlags(s4_1.flags() | Qt.ItemFlag.ItemIsEditable); self.asis_tree.setItemWidget(s4_1, 6, self.create_combo("수기"))
        s4_2 = QTreeWidgetItem(m4, ["4-2", "급여 명세서 생성 및 PDF 암호화 발송", "200", "0.1", "2", "40.0", "전산(NF)", "NF 시스템에서 명세서 생성 후 생년월일 암호화 걸어 발송"])
        s4_2.setFlags(s4_2.flags() | Qt.ItemFlag.ItemIsEditable); self.asis_tree.setItemWidget(s4_2, 6, self.create_combo("전산(NF)"))
        
        self.asis_tree.expandAll(); self.asis_tree.blockSignals(False); self.calc_totals()

    def run_ai(self):
        try: requests.get("http://localhost:11434/", timeout=2)
        except:
            QMessageBox.critical(self, "서버 연결 오류", "Ollama 서버가 꺼져 있습니다.")
            return

        data_list = []
        def traverse(item):
            row = []
            for i in range(8):
                if i == 6 and self.asis_tree.itemWidget(item, 6): row.append(self.asis_tree.itemWidget(item, 6).currentText())
                else: row.append(item.text(i))
            data_list.append(f"{'  ' * self.get_depth(item)}{'|'.join(row)}")
            for i in range(item.childCount()): traverse(item.child(i))

        for i in range(self.asis_tree.topLevelItemCount()): traverse(self.asis_tree.topLevelItem(i))
        if not data_list: return
        
        self.btn_run.setEnabled(False); self.btn_run.setText("⏳ AI 분석 진행 중... (최대 1~3분 소요)")
        self.progress.setVisible(True); self.progress.setRange(0, 0)
        self.analysis_text.setText("⏳ 로컬 LLM(Ollama)과 통신 중입니다...\n💡 AI가 인원(FTE) 최적화와 시스템 효율을 분석 중입니다.")

        self.worker = AIWorker("\n".join(data_list), self.model_combo.currentText(), self.proj_name.text())
        self.worker.finished.connect(self.on_success); self.worker.error.connect(self.on_fail)
        self.worker.start()

    def get_depth(self, item):
        d = 0
        while item.parent(): item = item.parent(); d += 1
        return d

    def on_success(self, rows, report):
        self.btn_run.setEnabled(True); self.btn_run.setText("✨ Gemma 3 12B 정밀 계층 분석 실행")
        self.progress.setVisible(False); self.analysis_text.setText(report)
        self.build_tree_from_data(self.tobe_tree, rows, is_asis=False)
        self.calc_totals() 

    def on_fail(self, msg):
        self.btn_run.setEnabled(True); self.btn_run.setText("✨ Gemma 3 12B 정밀 계층 분석 실행")
        self.progress.setVisible(False); self.analysis_text.setText("❌ 분석 실패")
        QMessageBox.critical(self, "오류", f"AI 분석 중 문제가 발생했습니다.\n{msg}")

    def export_csv(self):
        if not os.path.exists(self.secu_path): os.makedirs(self.secu_path)
        path = os.path.join(self.secu_path, f"ROI_Report_{datetime.now().strftime('%H%M%S')}.csv")
        try:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["구분", "No", "단위 프로세스명", "횟수(월)", "시간(hr)", "인원(명)", "총공수(hr)", "유형", "업무방법"])
                
                def write_tree(tree, prefix):
                    items = self.get_all_tree_items(tree)
                    for item in items:
                        row = [prefix]
                        for i in range(8):
                            if i == 6 and tree.itemWidget(item, 6): row.append(tree.itemWidget(item, 6).currentText())
                            else: row.append(item.text(i))
                        writer.writerow(row)
                
                write_tree(self.asis_tree, "AS-IS")
                writer.writerow([]) 
                write_tree(self.tobe_tree, "TO-BE")
            QMessageBox.information(self, "추출 성공", f"엑셀용 CSV 파일이 저장되었습니다.\n경로: {path}")
        except Exception as e:
            QMessageBox.critical(self, "추출 오류", str(e))

    def extract_tree_data(self, tree):
        data = []
        items = self.get_all_tree_items(tree)
        for item in items:
            row = []
            for i in range(8):
                if i == 6 and tree.itemWidget(item, 6): row.append(tree.itemWidget(item, 6).currentText())
                else: row.append(item.text(i))
            data.append(row)
        return data

    def build_tree_from_data(self, tree, rows, is_asis=True):
        tree.clear()
        item_map = {}
        for cols in rows:
            no = cols[0]
            parent_no = no.rsplit('-', 1)[0] if '-' in no else None
            
            if parent_no and parent_no in item_map: item = QTreeWidgetItem(item_map[parent_no], cols)
            else: item = QTreeWidgetItem(tree, cols)
            item_map[no] = item
            
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
            
            if cols[6] not in ["-", ""]: 
                tree.setItemWidget(item, 6, self.create_combo(cols[6], is_tobe=not is_asis))
            
            if not is_asis:
                if "자동화" in cols[6]: 
                    for i in range(8): item.setForeground(i, QColor("#9775fa")) 
                elif "전산" in cols[6]: 
                    for i in range(8): item.setForeground(i, QColor("#339af0")) 
        tree.expandAll()

    def save_data(self):
        try:
            if not os.path.exists(self.secu_path): os.makedirs(self.secu_path)
            path = os.path.join(self.secu_path, f"State_V44_{datetime.now().strftime('%H%M%S')}.json")
            full_state = {
                "metadata": {
                    "proj_name": self.proj_name.text(), "dept": self.dept.text(),
                    "author": self.author.text(), "wage": self.wage_input.text(), "model": self.model_combo.currentText()
                },
                "report": self.analysis_text.toPlainText(),
                "asis_data": self.extract_tree_data(self.asis_tree),
                "tobe_data": self.extract_tree_data(self.tobe_tree)
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(full_state, f, ensure_ascii=False, indent=4)
            QMessageBox.information(self, "저장 성공", "현재 작업 상태가 완벽하게 저장되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "저장 오류", str(e))

    def load_data(self):
        fname, _ = QFileDialog.getOpenFileName(self, "저장된 작업 파일 불러오기", self.secu_path, "JSON Files (*.json)")
        if fname:
            try:
                with open(fname, 'r', encoding='utf-8') as f: data = json.load(f)
                meta = data.get("metadata", {})
                self.proj_name.setText(meta.get("proj_name", "")); self.dept.setText(meta.get("dept", ""))
                self.author.setText(meta.get("author", "")); self.wage_input.setText(meta.get("wage", ""))
                self.model_combo.setCurrentText(meta.get("model", "gemma3:12b"))
                self.analysis_text.setText(data.get("report", ""))
                
                self.asis_tree.blockSignals(True)
                self.build_tree_from_data(self.asis_tree, data.get("asis_data", []), is_asis=True)
                self.asis_tree.blockSignals(False)
                
                self.tobe_tree.blockSignals(True)
                self.build_tree_from_data(self.tobe_tree, data.get("tobe_data", []), is_asis=False)
                self.tobe_tree.blockSignals(False)
                
                self.calc_totals() 
                QMessageBox.information(self, "불러오기 성공", "이전 작업 상태를 성공적으로 불러왔습니다.")
            except Exception as e:
                QMessageBox.critical(self, "불러오기 오류", str(e))

if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        app.setFont(QFont("Malgun Gothic", 10))
        app.setStyle("Fusion")
        win = ProcessAnalyzer()
        win.show()
        sys.exit(app.exec())
    except Exception as e:
        print("프로그램 실행 중 치명적 오류가 발생했습니다:")
        traceback.print_exc()
        input("엔터 키를 누르면 종료됩니다...")
