import os
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

EXCEL_PATH = 'Annual Report Database.xlsx'
TOC_IMAGE  = 'toc_background.png'


class GesiFullReportGenerator:
    def __init__(self, excel_path=EXCEL_PATH):
        self.doc = Document()
        self.xl = pd.ExcelFile(excel_path)

        style = self.doc.styles['Normal']
        style.font.name = '맑은 고딕'
        style.font.size = Pt(11)

    def _read_sheet(self, sheet_name, **kwargs):
        return pd.read_excel(self.xl, sheet_name=sheet_name, **kwargs)

    # ------------------------------------------------------------------ #
    #  1. 표지                                                              #
    # ------------------------------------------------------------------ #
    def add_title_page(self):
        for _ in range(5):
            self.doc.add_paragraph()

        title = self.doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run("2025 GESI ANNUAL REPORT")
        run.bold = True
        run.font.size = Pt(32)
        run.font.color.rgb = RGBColor(34, 139, 34)

        subtitle = self.doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run2 = subtitle.add_run("탄소중립을 향한 데이터와 현장의 기록")
        run2.font.size = Pt(16)

        self.doc.add_page_break()

    # ------------------------------------------------------------------ #
    #  2. 목차 (이미지 포함)                                                #
    # ------------------------------------------------------------------ #
    def add_toc_page(self, image_path=TOC_IMAGE):
        """목차 페이지 — 상단에 비전 이미지, 하단에 목차 항목"""

        # 상단 이미지
        if image_path and os.path.exists(image_path):
            p_img = self.doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_img = p_img.add_run()
            run_img.add_picture(image_path, width=Inches(6.2))
        else:
            self.doc.add_paragraph()

        self.doc.add_paragraph()

        # 목차 제목
        toc_title = self.doc.add_paragraph()
        toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = toc_title.add_run("목  차  /  CONTENTS")
        r.bold = True
        r.font.size = Pt(16)
        r.font.color.rgb = RGBColor(34, 139, 34)

        self.doc.add_paragraph()

        toc_items = [
            ("01", "기관 소개"),
            ("02", "연구 연혁 (2015–2025)"),
            ("03", "2025 핵심 연구 성과"),
            ("04", "주요 활동 및 외부 소통"),
            ("05", "2025년 주요 연구 상세"),
        ]

        for num, item_title in toc_items:
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.8)
            r_num = p.add_run(f"{num}.  ")
            r_num.bold = True
            r_num.font.size = Pt(12)
            r_num.font.color.rgb = RGBColor(34, 139, 34)
            r_body = p.add_run(item_title)
            r_body.font.size = Pt(12)

            # Pillar 소항목 (05번)
            if num == "05":
                try:
                    df_2025 = self._read_sheet('2025', index_col=0)
                    for col in df_2025.columns:
                        p_sub = self.doc.add_paragraph()
                        p_sub.paragraph_format.left_indent = Inches(1.4)
                        r_sub = p_sub.add_run(f"∙  {col}")
                        r_sub.font.size = Pt(10)
                        r_sub.font.color.rgb = RGBColor(90, 90, 90)
                except Exception:
                    pass

        self.doc.add_page_break()

    # ------------------------------------------------------------------ #
    #  3. 기관 소개 (Institute_Info)                                        #
    # ------------------------------------------------------------------ #
    def add_institute_intro(self):
        self.doc.add_heading('01. 기관 소개', level=1)
        try:
            df = self._read_sheet('Institute_Info')
            row_2025 = df[df.iloc[:, 0] == 2025]
            row = row_2025.iloc[0] if not row_2025.empty else df.iloc[0]

            col_overview = df.columns[1]
            col_people   = df.columns[2]
            col_contact  = df.columns[3]

            if pd.notna(row[col_overview]):
                p = self.doc.add_paragraph(str(row[col_overview]))
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            self.doc.add_paragraph()

            tbl = self.doc.add_table(rows=2, cols=2)
            tbl.style = 'Light Grid Accent 3'
            tbl.rows[0].cells[0].text = '기관 인력'
            tbl.rows[0].cells[1].text = str(row[col_people])  if pd.notna(row[col_people])  else '-'
            tbl.rows[1].cells[0].text = '주요 연락처'
            tbl.rows[1].cells[1].text = str(row[col_contact]) if pd.notna(row[col_contact]) else '-'
        except Exception as e:
            self.doc.add_paragraph(f"기관 소개 데이터를 불러올 수 없습니다. ({e})")

        self.doc.add_paragraph()
        self.doc.add_page_break()

    # ------------------------------------------------------------------ #
    #  4. 연구 연혁 (Research_History) — 인포그래픽 포함                   #
    # ------------------------------------------------------------------ #
    def _create_timeline_infographic(self) -> str:
        """연구 연혁 타임라인 인포그래픽을 PNG로 저장하고 경로를 반환"""
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches

        # Windows 한글 폰트
        for fname in ('Malgun Gothic', 'NanumGothic', 'AppleGothic', 'sans-serif'):
            try:
                matplotlib.rc('font', family=fname)
                break
            except Exception:
                continue
        matplotlib.rc('axes', unicode_minus=False)

        df_hist   = self._read_sheet('Research_History')
        df_master = self._read_sheet('Master_Research')

        year_col    = df_hist.columns[0]
        pillar_cols = df_hist.columns[1:]

        df_hist = df_hist.copy()
        df_hist[year_col] = df_hist[year_col].ffill()

        # (year, pillar_idx) → 연구 개수
        yp_count: dict = {}
        for _, row in df_hist.iterrows():
            yr = int(row[year_col])
            for p_idx, p_col in enumerate(pillar_cols):
                if pd.notna(row[p_col]):
                    yp_count[(yr, p_idx)] = yp_count.get((yr, p_idx), 0) + 1

        # Master_Research ID → (year, pillar_idx)
        col_id     = df_master.columns[0]
        col_parent = df_master.columns[1]
        id_to_pos: dict = {}
        for _, row in df_master.iterrows():
            rid = str(row[col_id]).strip()
            parts = rid.split('-')
            if len(parts) == 4:
                try:
                    id_to_pos[rid] = (int(parts[1]), int(parts[2]) - 1)
                except ValueError:
                    pass

        connections = []
        for _, row in df_master.iterrows():
            rid = str(row[col_id]).strip()
            parent = row[col_parent]
            if pd.notna(parent):
                pid = str(parent).strip()
                if rid in id_to_pos and pid in id_to_pos:
                    connections.append((id_to_pos[pid], id_to_pos[rid]))

        # ---- Figure ----
        fig, ax = plt.subplots(figsize=(16, 6.5))
        fig.patch.set_facecolor('#f0f4f0')
        ax.set_facecolor('#f0f4f0')

        PILLAR_COLORS = ['#2e7d32', '#1565c0', '#bf360c']
        PILLAR_Y      = [4.0, 2.0, 0.0]
        PILLAR_LABELS = [
            'Pillar 1\n시스템 전환\n& 섹터커플링',
            'Pillar 2\n지역 에너지 전환\n& 분산에너지',
            'Pillar 3\n탈탄소 정책\n& 사회적 가치',
        ]
        YEARS = list(range(2015, 2026))

        # 레인 배경 + 레이블
        for p_idx in range(3):
            y = PILLAR_Y[p_idx]
            rect = mpatches.FancyBboxPatch(
                (2014.7, y - 0.72), 10.8, 1.44,
                boxstyle='round,pad=0.05', linewidth=0,
                facecolor=PILLAR_COLORS[p_idx], alpha=0.10
            )
            ax.add_patch(rect)
            ax.text(2014.35, y, PILLAR_LABELS[p_idx],
                    ha='right', va='center', fontsize=8,
                    color=PILLAR_COLORS[p_idx], fontweight='bold', linespacing=1.5)

        # 연도 수직선 + 라벨
        for yr in YEARS:
            ax.axvline(x=yr, color='#bbbbbb', linewidth=0.6,
                       linestyle=':', alpha=0.9, zorder=1)
            ax.text(yr, -1.1, str(yr), ha='center', va='top',
                    fontsize=7.5, color='#666666')

        # 연구 버블
        for (yr, p_idx), cnt in yp_count.items():
            y_base = PILLAR_Y[p_idx]
            # 버블 크기: 연구 수에 비례
            size = 60 + cnt * 20
            ax.scatter(yr, y_base, s=min(size, 220),
                       color=PILLAR_COLORS[p_idx], alpha=0.75,
                       zorder=5, edgecolors='white', linewidth=1.5)
            if cnt > 1:
                ax.text(yr, y_base, str(cnt), ha='center', va='center',
                        fontsize=6.5, color='white', fontweight='bold', zorder=6)

        # 연결 화살표 (연구 분야 ID 기반)
        for (from_yr, from_p), (to_yr, to_p) in connections:
            fy, ty = PILLAR_Y[from_p], PILLAR_Y[to_p]
            rad = 0.35 if from_p == to_p else -0.45
            ax.annotate('',
                xy=(to_yr, ty), xytext=(from_yr, fy),
                arrowprops=dict(
                    arrowstyle='->', color='#e65100', lw=2.2,
                    connectionstyle=f'arc3,rad={rad}',
                    shrinkA=9, shrinkB=9
                ), zorder=7)

        # 가로 흐름 화살표 (각 필라 레인 우측 끝)
        for p_idx in range(3):
            y = PILLAR_Y[p_idx]
            ax.annotate('', xy=(2025.6, y), xytext=(2014.8, y),
                        arrowprops=dict(
                            arrowstyle='->', color=PILLAR_COLORS[p_idx],
                            lw=1.2, alpha=0.35
                        ))

        ax.set_title('GESI 연구 연혁 타임라인  (2015–2025)',
                     fontsize=14, fontweight='bold', color='#1a3d2b', pad=14)

        legend_handles = [
            mpatches.Patch(color=PILLAR_COLORS[0], alpha=0.8,
                           label='Pillar 1: 시스템 전환 & 섹터커플링'),
            mpatches.Patch(color=PILLAR_COLORS[1], alpha=0.8,
                           label='Pillar 2: 지역 에너지 전환 & 분산에너지'),
            mpatches.Patch(color=PILLAR_COLORS[2], alpha=0.8,
                           label='Pillar 3: 탈탄소 정책 & 사회적 가치'),
            mpatches.Patch(color='#e65100', alpha=0.8,
                           label='연구 연계 / 확장 (연구 분야 ID 기반)'),
        ]
        ax.legend(handles=legend_handles, loc='upper right',
                  fontsize=8.5, framealpha=0.9, edgecolor='#cccccc')

        # 버블 크기 범례 (주석)
        ax.text(2015, -1.75,
                '※ 버블 크기 = 해당 연도·분야 연구 수  |  숫자 = 연구 건수',
                fontsize=7.5, color='#888888')

        ax.set_xlim(2013.8, 2026.2)
        ax.set_ylim(-2.1, 5.5)
        ax.axis('off')
        plt.tight_layout(pad=1.5)

        out_path = 'research_timeline.png'
        plt.savefig(out_path, dpi=150, bbox_inches='tight',
                    facecolor=fig.get_facecolor())
        plt.close()
        return out_path

    def add_research_history(self):
        """연구 연혁 섹션 (Research_History 시트 + 타임라인 인포그래픽)"""
        self.doc.add_heading('02. 연구 연혁 (2015–2025)', level=1)
        self.doc.add_paragraph(
            "GESI가 걸어온 지난 10년은 한국 에너지 전환의 역사와 궤를 같이 합니다. "
            "아래 타임라인은 세 가지 핵심 연구 축이 어떻게 심화·확장되어 왔는지를 보여줍니다."
        )
        self.doc.add_paragraph()

        # 인포그래픽 삽입
        try:
            infographic_path = self._create_timeline_infographic()
            p_img = self.doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run().add_picture(infographic_path, width=Inches(6.4))
        except Exception as e:
            self.doc.add_paragraph(f"[인포그래픽 생성 오류: {e}]")

        self.doc.add_paragraph()

        # 연도별 상세 표
        try:
            df = self._read_sheet('Research_History')
            year_col    = df.columns[0]
            pillar_cols = df.columns[1:]
            df = df.copy()
            df[year_col] = df[year_col].ffill()

            n_cols = 1 + len(pillar_cols)
            tbl = self.doc.add_table(rows=1, cols=n_cols)
            tbl.style = 'Light Grid Accent 3'
            hdr = tbl.rows[0].cells
            hdr[0].text = '연도'
            for i, col in enumerate(pillar_cols):
                hdr[i + 1].text = str(col)

            for _, row in df.iterrows():
                cells = tbl.add_row().cells
                yr_val = row[year_col]
                cells[0].text = str(int(yr_val)) if pd.notna(yr_val) else ''
                for i, col in enumerate(pillar_cols):
                    cells[i + 1].text = str(row[col]) if pd.notna(row[col]) else ''
        except Exception as e:
            self.doc.add_paragraph(f"연혁 데이터를 불러올 수 없습니다. ({e})")

        self.doc.add_page_break()

    # ------------------------------------------------------------------ #
    #  5. 2025 핵심 연구 성과 (Master_Research)                            #
    # ------------------------------------------------------------------ #
    def add_key_research_2025(self):
        self.doc.add_heading('03. 2025 핵심 연구 성과', level=1)
        try:
            df = self._read_sheet('Master_Research')
            col_field   = df.columns[3]
            col_title   = df.columns[4]
            col_summary = df.columns[5]
            col_output  = df.columns[7]

            pillar_intros = {
                '1': "에너지 시스템 통합과 섹터커플링을 통한 유연성 확보 연구입니다.",
                '2': "지방정부 주도의 에너지 분권과 거버넌스 강화 연구입니다.",
                '3': "탈탄소 전략과 보조금 개편 등 사회적 가치 실현 연구입니다.",
            }

            for field_val, group in df.groupby(col_field):
                self.doc.add_heading(str(field_val), level=2)
                for key, intro in pillar_intros.items():
                    if key in str(field_val):
                        self.doc.add_paragraph(intro)
                        break

                for _, row in group.iterrows():
                    p = self.doc.add_paragraph()
                    p.add_run(f"▣ {row[col_title]}").bold = True
                    if pd.notna(row[col_summary]):
                        self.doc.add_paragraph(f"  - 핵심 요약: {row[col_summary]}")
                    if pd.notna(row[col_output]):
                        self.doc.add_paragraph(f"  - 성과물: {row[col_output]}")
        except Exception as e:
            self.doc.add_paragraph(f"연구 데이터를 불러올 수 없습니다. ({e})")

    # ------------------------------------------------------------------ #
    #  6. 주요 활동 및 소식 (Activities_&_News)                            #
    # ------------------------------------------------------------------ #
    def add_activities(self):
        self.doc.add_page_break()
        self.doc.add_heading('04. 주요 활동 및 외부 소통', level=1)
        self.doc.add_paragraph(
            "GESI는 연구를 넘어 시민사회 및 정책 입안자들과 끊임없이 소통하고 있습니다."
        )
        try:
            df = self._read_sheet('Activities_&_News')
            col_date    = df.columns[1]
            col_type    = df.columns[3]
            col_name    = df.columns[4]
            col_content = df.columns[6]

            for _, row in df.iterrows():
                date_str = str(row[col_date]) if pd.notna(row[col_date]) else ''
                act_name = str(row[col_name]) if pd.notna(row[col_name]) else ''
                act_type = str(row[col_type]) if pd.notna(row[col_type]) else ''
                p = self.doc.add_paragraph()
                p.add_run(f"• [{date_str}] {act_name}  ({act_type})").bold = True
                if pd.notna(row[col_content]):
                    self.doc.add_paragraph(f"  : {row[col_content]}")
        except Exception as e:
            self.doc.add_paragraph(f"활동 데이터를 불러올 수 없습니다. ({e})")

    # ------------------------------------------------------------------ #
    #  7. 2025 Pillar별 한 페이지씩 (2025 시트)                           #
    # ------------------------------------------------------------------ #
    def add_2025_pillar_pages(self):
        self.doc.add_page_break()
        self.doc.add_heading('05. 2025년 주요 연구 상세', level=1)
        try:
            df = self._read_sheet('2025', index_col=0)
            for pillar_name in df.columns:
                heading = self.doc.add_heading('', level=2)
                run = heading.add_run(str(pillar_name))
                run.font.color.rgb = RGBColor(34, 139, 34)

                for row_label, cell_value in df[pillar_name].items():
                    if pd.isna(cell_value):
                        continue
                    label_para = self.doc.add_paragraph()
                    label_para.add_run(str(row_label).strip()).bold = True
                    content_para = self.doc.add_paragraph(
                        str(cell_value).strip().strip('"')
                    )
                    content_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    self.doc.add_paragraph()

                self.doc.add_page_break()
        except Exception as e:
            self.doc.add_paragraph(f"2025 Pillar 데이터를 불러올 수 없습니다. ({e})")

    # ------------------------------------------------------------------ #
    #  저장                                                                #
    # ------------------------------------------------------------------ #
    def save_report(self, filename):
        self.doc.save(filename)
        print(f"보고서 생성 완료: {filename}")


# --- 실행부 ---
if __name__ == "__main__":
    gen = GesiFullReportGenerator(EXCEL_PATH)
    gen.add_title_page()
    gen.add_toc_page(TOC_IMAGE)
    gen.add_institute_intro()
    gen.add_research_history()
    gen.add_key_research_2025()
    gen.add_activities()
    gen.add_2025_pillar_pages()
    gen.save_report("GESI_2025_Annual_Report_Final.docx")
