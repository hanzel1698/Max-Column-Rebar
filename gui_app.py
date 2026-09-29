import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import queue
import sys
import io
import os
from datetime import datetime
import pythoncom
import comtypes.client
from comtypes import IUnknown, POINTER
import ctypes
import re
import pandas as pd
import numpy as np

# ReportLab imports for generating a premium PDF
from reportlab.lib.pagesizes import letter, A4, A3, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
from reportlab.pdfgen import canvas
from reportlab.platypus import Flowable

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class NumberedCanvas(canvas.Canvas):
    """
    Custom canvas to enable 2-pass page numbering and professional header/footers 
    on A3 landscape pages.
    """
    model_name = "Untitled"
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
        
    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()
        
    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            super().showPage()
        super().save()
        
    def draw_page_number(self, page_count):
        # Page 1 is the cover page - do not draw header/footer
        if self._pageNumber == 1:
            return
            
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#718096"))
        
        # Header (A3 Landscape dimensions: 1190.55 x 841.89)
        self.drawString(54, 796, "ETABS COLUMN REBAR ANALYSIS & GRID KEY PLAN")
        self.drawRightString(1136, 796, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.setStrokeColor(colors.HexColor("#E2E8F0"))
        self.setLineWidth(0.5)
        self.line(54, 788, 1136, 788)
        
        # Footer
        self.line(54, 45, 1136, 45)
        self.drawString(54, 30, f"Model: {self.model_name} | Confidential - Engineering Design Document")
        page_text = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(1136, 30, page_text)
        self.restoreState()


def rects_intersect(r1, r2, margin=1.0):
    """
    Helper to check if two bounding box rectangles intersect.
    Rectangles are tuples in format (x_min, y_min, x_max, y_max)
    """
    return not (r1[2] + margin < r2[0] or 
                r1[0] - margin > r2[2] or 
                r1[3] + margin < r2[1] or 
                r1[1] - margin > r2[3])


class ColumnPlanFlowable(Flowable):
    """
    Custom Flowable to draw the scaled column layout plan on a spacious A3 canvas,
    utilizing an AABB collision-avoidance label placement algorithm.
    """
    def __init__(self, x_grids, y_grids, aggregated_grids, width, height, scale_factor=1.5, low_threshold=2.7, high_threshold=3.0):
        super().__init__()
        self.x_grids = x_grids
        self.y_grids = y_grids
        self.aggregated_grids = aggregated_grids
        self.width = width
        self.height = height
        self.scale_factor = scale_factor
        self.low_threshold = low_threshold
        self.high_threshold = high_threshold
        
    def wrap(self, availWidth, availHeight):
        return self.width, self.height
        
    def draw(self):
        canvas = self.canv
        
        # Gather all X and Y coordinates to determine bounding box
        x_coords = [val for _, val in self.x_grids]
        y_coords = [val for _, val in self.y_grids]
        
        for col in self.aggregated_grids:
            x_coords.append(col['x'])
            y_coords.append(col['y'])
            
        if not x_coords or not y_coords:
            canvas.setFont("Helvetica", 10)
            canvas.drawCentredString(self.width/2, self.height/2, "No columns to draw")
            return
            
        x_min, x_max = min(x_coords), max(x_coords)
        y_min, y_max = min(y_coords), max(y_coords)
        
        span_x = x_max - x_min
        span_y = y_max - y_min
        
        if span_x == 0: span_x = 1.0
        if span_y == 0: span_y = 1.0
        
        # Add 12% padding around coordinates to accommodate grid bubbles
        x_min_padded = x_min - 0.12 * span_x
        x_max_padded = x_max + 0.12 * span_x
        y_min_padded = y_min - 0.12 * span_y
        y_max_padded = y_max + 0.12 * span_y
        
        padded_span_x = x_max_padded - x_min_padded
        padded_span_y = y_max_padded - y_min_padded
        
        # Calculate scale to fit plan inside flowable dimensions
        scale = min(self.width / padded_span_x, self.height / padded_span_y)
        
        # Center plan
        x_offset = (self.width - scale * padded_span_x) / 2 - scale * x_min_padded
        y_offset = (self.height - scale * padded_span_y) / 2 - scale * y_min_padded
        
        def to_pdf(x, y):
            return x_offset + scale * x, y_offset + scale * y
            
        # Draw physical boundary
        canvas.setStrokeColor(colors.HexColor("#EDF2F7"))
        canvas.setLineWidth(1)
        canvas.setFillColor(colors.HexColor("#F8FAFC"))
        canvas.roundRect(0, 0, self.width, self.height, 8, stroke=1, fill=1)
        
        # Define limits for drawing grid lines
        grid_y_start = y_min - 0.04 * span_y
        grid_y_end = y_max + 0.04 * span_y
        grid_x_start = x_min - 0.04 * span_x
        grid_x_end = x_max + 0.04 * span_x
        
        # 1. Draw X Grid Lines (vertical lines with a constant X coordinate)
        bubble_r = 8 * self.scale_factor
        canvas.setFont("Helvetica-Bold", 7.0 * self.scale_factor)
        
        for label, gx in self.x_grids:
            x_pdf, y_start_pdf = to_pdf(gx, grid_y_start)
            x_pdf, y_end_pdf = to_pdf(gx, grid_y_end)
            
            # Dashed grid line
            canvas.setStrokeColor(colors.HexColor("#CBD5E1"))
            canvas.setLineWidth(0.5)
            canvas.setDash(4, 4)
            canvas.line(x_pdf, y_start_pdf, x_pdf, y_end_pdf)
            
            # Solid bubble at bottom
            canvas.setStrokeColor(colors.HexColor("#475569"))
            canvas.setLineWidth(0.7)
            canvas.setDash()
            canvas.setFillColor(colors.HexColor("#E2E8F0"))
            canvas.circle(x_pdf, y_start_pdf - bubble_r, bubble_r, stroke=1, fill=1)
            # Text in bubble
            canvas.setFillColor(colors.HexColor("#1E293B"))
            canvas.drawCentredString(x_pdf, y_start_pdf - bubble_r - (2.5 * self.scale_factor), label)
            
            # Solid bubble at top
            canvas.setFillColor(colors.HexColor("#E2E8F0"))
            canvas.circle(x_pdf, y_end_pdf + bubble_r, bubble_r, stroke=1, fill=1)
            canvas.setFillColor(colors.HexColor("#1E293B"))
            canvas.drawCentredString(x_pdf, y_end_pdf + bubble_r - (2.5 * self.scale_factor), label)
            
        # 2. Draw Y Grid Lines (horizontal lines with a constant Y coordinate)
        for label, gy in self.y_grids:
            x_start_pdf, y_pdf = to_pdf(grid_x_start, gy)
            x_end_pdf, y_pdf = to_pdf(grid_x_end, gy)
            
            # Dashed grid line
            canvas.setStrokeColor(colors.HexColor("#CBD5E1"))
            canvas.setLineWidth(0.5)
            canvas.setDash(4, 4)
            canvas.line(x_start_pdf, y_pdf, x_end_pdf, y_pdf)
            
            # Solid bubble at left
            canvas.setStrokeColor(colors.HexColor("#475569"))
            canvas.setLineWidth(0.7)
            canvas.setDash()
            canvas.setFillColor(colors.HexColor("#E2E8F0"))
            canvas.circle(x_start_pdf - bubble_r, y_pdf, bubble_r, stroke=1, fill=1)
            canvas.setFillColor(colors.HexColor("#1E293B"))
            canvas.drawCentredString(x_start_pdf - bubble_r, y_pdf - (2.5 * self.scale_factor), label)
            
            # Solid bubble at right
            canvas.setFillColor(colors.HexColor("#E2E8F0"))
            canvas.circle(x_end_pdf + bubble_r, y_pdf, bubble_r, stroke=1, fill=1)
            canvas.setFillColor(colors.HexColor("#1E293B"))
            canvas.drawCentredString(x_end_pdf + bubble_r, y_pdf - (2.5 * self.scale_factor), label)
            
        # Column scaling factor
        col_scale = scale
        avg_col_size_m = 0.45
        avg_col_size_pt = scale * avg_col_size_m
        
        # Calculate minimum distance between any two distinct columns in meters
        min_dist_m = float('inf')
        for i in range(len(self.aggregated_grids)):
            c1 = self.aggregated_grids[i]
            for j in range(i + 1, len(self.aggregated_grids)):
                c2 = self.aggregated_grids[j]
                dx = c1['x'] - c2['x']
                dy = c1['y'] - c2['y']
                dist = (dx**2 + dy**2)**0.5
                if dist > 0.05 and dist < min_dist_m:
                    min_dist_m = dist
                    
        # Determine the maximum allowed column drawing size based on spacing and limits
        max_col_size_pt = 4.0 * self.scale_factor
        if min_dist_m != float('inf'):
            min_dist_pt = min_dist_m * scale
            max_col_size_pt = min(4.0 * self.scale_factor, max(1.8 * self.scale_factor, min_dist_pt * 0.7 * self.scale_factor))
            
        min_col_size_pt = min(1.8 * self.scale_factor, max_col_size_pt)
        
        # Enforce bounds
        if avg_col_size_pt < min_col_size_pt:
            col_scale = scale * (min_col_size_pt / avg_col_size_pt)
            
        if col_scale * avg_col_size_m > max_col_size_pt:
            col_scale = max_col_size_pt / avg_col_size_m
            
        occupied_rects = []
        
        # Bounding box calculation for columns
        for col in self.aggregated_grids:
            cx, cy = to_pdf(col['x'], col['y'])
            w = col['width'] * col_scale
            d = col['depth'] * col_scale
            r = max(w, d) / 2.0
            col_rect = (cx - r - 1.5, cy - r - 1.5, cx + r + 1.5, cy + r + 1.5)
            occupied_rects.append(col_rect)
            
        # 3. Draw Column Shapes
        for col in self.aggregated_grids:
            cx, cy = to_pdf(col['x'], col['y'])
            w = col['width'] * col_scale
            d = col['depth'] * col_scale
            shape = col['shape']
            angle = col['angle']
            
            canvas.saveState()
            canvas.translate(cx, cy)
            canvas.rotate(angle)
            
            canvas.setFillColor(colors.HexColor("#1E293B"))
            canvas.setStrokeColor(colors.HexColor("#0F172A"))
            canvas.setLineWidth(0.8)
            canvas.setDash()
            
            if shape == "Circular":
                r = w / 2.0
                canvas.circle(0, 0, r, stroke=1, fill=1)
                canvas.setStrokeColor(colors.HexColor("#475569"))
                canvas.setLineWidth(0.4)
                canvas.line(-r, 0, r, 0)
                canvas.line(0, -r, 0, r)
            else:
                canvas.rect(-d/2, -w/2, d, w, stroke=1, fill=1)
                canvas.setStrokeColor(colors.HexColor("#475569"))
                canvas.setLineWidth(0.4)
                canvas.line(-d/2, -w/2, d/2, w/2)
                canvas.line(-d/2, w/2, d/2, -w/2)
                
            canvas.restoreState()
            
        # 4. Draw Rebar % Badges with Collision Avoidance
        for col in self.aggregated_grids:
            cx, cy = to_pdf(col['x'], col['y'])
            w = col['width'] * col_scale
            d = col['depth'] * col_scale
            
            rebar = col['max_rebar']
            
            if rebar == 0.0:
                bg_color = colors.HexColor("#64748B")
                text_str = "N/D"
            elif rebar < self.low_threshold:
                bg_color = colors.HexColor("#10B981")
                text_str = f"{rebar:.2f}%"
            elif rebar <= self.high_threshold:
                bg_color = colors.HexColor("#F59E0B")
                text_str = f"{rebar:.2f}%"
            else:
                bg_color = colors.HexColor("#EF4444")
                text_str = f"{rebar:.2f}%"
                
            font_sz = 6.5 * self.scale_factor
            canvas.setFont("Helvetica-Bold", font_sz)
            text_width = canvas.stringWidth(text_str, "Helvetica-Bold", font_sz)
            badge_w = text_width + 4 * self.scale_factor
            badge_h = 9 * self.scale_factor
            
            r_col = max(w, d) / 2.0
            
            candidates = [
                (r_col + 2, r_col + 2),
                (-r_col - badge_w - 2, -r_col - badge_h - 2),
                (-r_col - badge_w - 2, r_col + 2),
                (r_col + 2, -r_col - badge_h - 2),
                (-badge_w/2, r_col + 3),
                (-badge_w/2, -r_col - badge_h - 3),
                (r_col + 3, -badge_h/2),
                (-r_col - badge_w - 3, -badge_h/2)
            ]
            
            selected_pos = None
            for dx, dy in candidates:
                bx = cx + dx
                by = cy + dy
                candidate_rect = (bx, by, bx + badge_w, by + badge_h)
                
                has_collision = False
                for r_occ in occupied_rects:
                    if rects_intersect(candidate_rect, r_occ, margin=0.5):
                        has_collision = True
                        break
                        
                if not has_collision:
                    selected_pos = (bx, by, candidate_rect)
                    break
                    
            if not selected_pos:
                bx = cx + r_col + 4
                by = cy + r_col + 4
                candidate_rect = (bx, by, bx + badge_w, by + badge_h)
                selected_pos = (bx, by, candidate_rect)
                
            bx, by, final_rect = selected_pos
            occupied_rects.append(final_rect)
            
            # Draw badge background
            canvas.setFillColor(bg_color)
            canvas.roundRect(bx, by, badge_w, badge_h, 2, stroke=0, fill=1)
            
            # Print badge text
            canvas.setFillColor(colors.white)
            canvas.drawString(bx + (2 * self.scale_factor), by + ((badge_h - font_sz)/2.0 + 0.5), text_str)
            
            # Add an invisible AcroForm text field as a hover tooltip for Governing Storey
            try:
                form = canvas.acroForm
                tooltip_text = f"{col.get('max_rebar_story', 'N/A')}"
                field_name = f"tip_{col.get('grid_label', 'NA')}_{col.get('max_rebar_frame', 'NA')}"
                form.textfield(
                    name=field_name,
                    tooltip=tooltip_text,
                    x=bx, y=by,
                    width=badge_w, height=badge_h,
                    borderStyle='solid', forceBorder=False,
                    fillColor=colors.transparent, borderColor=colors.transparent, textColor=colors.transparent
                )
            except Exception:
                pass
            
            # Print Column Section Size label centered below the column (e.g. 40X50)
            canvas.setFillColor(colors.HexColor("#475569"))
            canvas.setFont("Helvetica-Bold", 6.0 * self.scale_factor)
            lbl_offset = r_col + 6 * self.scale_factor
            if col['shape'] == "Circular":
                size_str = f"Dia {col['width']*100:.0f}"
            else:
                size_str = f"{col['width']*100:.0f}X{col['depth']*100:.0f}"
            canvas.drawCentredString(cx, cy - lbl_offset, size_str)
            
        # 5. Draw Legend Box at bottom right of the canvas
        legend_x = self.width - 150 * self.scale_factor
        legend_y = 15
        legend_w = 135 * self.scale_factor
        legend_h = 65 * self.scale_factor
        
        canvas.setStrokeColor(colors.HexColor("#CBD5E1"))
        canvas.setLineWidth(0.5)
        canvas.setFillColor(colors.HexColor("#FFFFFF"))
        canvas.roundRect(legend_x, legend_y, legend_w, legend_h, 4, stroke=1, fill=1)
        
        canvas.setFont("Helvetica-Bold", 7.5 * self.scale_factor)
        canvas.setFillColor(colors.HexColor("#1E293B"))
        canvas.drawString(legend_x + 8 * self.scale_factor, legend_y + legend_h - (10 * self.scale_factor), "LEGEND: Rebar %")
        
        entries = [
            (colors.HexColor("#10B981"), f"Rebar < {self.low_threshold:.2f}% (Low)"),
            (colors.HexColor("#F59E0B"), f"{self.low_threshold:.2f}% <= Rebar <= {self.high_threshold:.2f}%"),
            (colors.HexColor("#EF4444"), f"Rebar > {self.high_threshold:.2f}% (High)"),
            (colors.HexColor("#64748B"), "N/D (Not Designed)")
        ]
        
        for idx, (color, label_text) in enumerate(entries):
            ey = legend_y + legend_h - (22 * self.scale_factor) - idx * (10.5 * self.scale_factor)
            canvas.setFillColor(color)
            canvas.roundRect(legend_x + 8 * self.scale_factor, ey, 20 * self.scale_factor, 6.5 * self.scale_factor, 1.5, stroke=0, fill=1)
            canvas.setFillColor(colors.HexColor("#475569"))
            canvas.setFont("Helvetica", 6.0 * self.scale_factor)
            canvas.drawString(legend_x + 32 * self.scale_factor, ey + (1.0 * self.scale_factor), label_text)
            
        # Draw physical scale indicator at bottom left
        scale_x = 15
        scale_y = 15
        scale_w_pt = 60
        scale_m = scale_w_pt / scale
        canvas.setStrokeColor(colors.HexColor("#475569"))
        canvas.setLineWidth(1)
        canvas.line(scale_x, scale_y, scale_x + scale_w_pt, scale_y)
        canvas.line(scale_x, scale_y - 2, scale_x, scale_y + 2)
        canvas.line(scale_x + scale_w_pt, scale_y - 2, scale_x + scale_w_pt, scale_y + 2)
        canvas.setFont("Helvetica", 6.0 * self.scale_factor)
        canvas.setFillColor(colors.HexColor("#475569"))
        canvas.drawString(scale_x, scale_y + (4 * self.scale_factor), f"Scale Bar: {scale_m:.1f} m")


def get_section_dimensions(sap_model, prop_name):
    """Retrieve physical widths/depths from ETABS property naming definitions."""
    try:
        ret = sap_model.PropFrame.GetRectangle(prop_name)
        if ret[-1] == 0:
            return "Rectangular", ret[3], ret[2]
    except Exception:
        pass
    try:
        ret = sap_model.PropFrame.GetCircle(prop_name)
        if ret[-1] == 0:
            return "Circular", ret[2], ret[2]
    except Exception:
        pass
    return "Rectangular", 0.4, 0.4


def read_etabs_table(sap_model, table_key):
    """Query a specific ETABS database table to raw pandas DataFrame structure."""
    try:
        ret = sap_model.DatabaseTables.GetTableForDisplayCSVString(TableKey=table_key, GroupName="")
        if isinstance(ret, str):
            if "Story" in ret or "Label" in ret or "," in ret:
                return pd.read_csv(io.StringIO(ret))
        if isinstance(ret, (tuple, list)):
            for val in ret:
                if isinstance(val, str) and ("," in val or "\n" in val or "Frame" in val or "Story" in val or "Grid" in val):
                    return pd.read_csv(io.StringIO(val))
    except Exception:
        pass
    return None


def get_grid_lines(sap_model):
    """Extract coordinate grid lines from ETABS definitions."""
    table_key = "Grid Definitions - Grid Lines"
    df_results = read_etabs_table(sap_model, table_key)
    if df_results is None:
        return None, None
    try:
        df_results.columns = [c.strip() for c in df_results.columns]
        x_grids, y_grids = [], []
        for _, row in df_results.iterrows():
            g_type = str(row['Grid Line Type']).strip()
            g_id = str(row['ID']).strip()
            try:
                ord_val = float(row['Ordinate'])
            except ValueError:
                continue
            if 'X' in g_type:
                x_grids.append((g_id, ord_val))
            elif 'Y' in g_type:
                y_grids.append((g_id, ord_val))
        return sorted(list(set(x_grids)), key=lambda x: x[1]), sorted(list(set(y_grids)), key=lambda x: x[1])
    except Exception:
        return None, None


def extract_rebar_design_results(sap_model):
    """Extract rebar percentage ratios from active ETABS models."""
    table_keys = []
    try:
        ret = sap_model.DatabaseTables.GetAvailableTables()
        num_tables, keys, names, import_types, ret_code = ret
        if ret_code == 0:
            table_keys = list(keys)
    except Exception:
        pass
        
    design_table_key = None
    for key in table_keys:
        k_lower = key.lower()
        if "concrete column" in k_lower and "pmm" in k_lower and "envelope" in k_lower:
            design_table_key = key
            break
            
    common_keys = [
        "Concrete Column PMM Envelope - IS 456-2000",
        "Concrete Column PMM Envelope - ACI 318-19",
        "Concrete Column PMM Envelope - ACI 318-14",
        "Concrete Column PMM Envelope - Eurocode 2-2004",
        "Concrete Column PMM Envelope"
    ]
    
    df_results = None
    if design_table_key:
        df_results = read_etabs_table(sap_model, design_table_key)
    else:
        for key in common_keys:
            df_results = read_etabs_table(sap_model, key)
            if df_results is not None:
                break
                
    rebar_data = {}
    if df_results is not None:
        df_results.columns = [c.strip() for c in df_results.columns]
        frame_col = None
        for target in ['uniquename', 'unique name', 'frame', 'element']:
            for c in df_results.columns:
                if c.lower() == target:
                    frame_col = c
                    break
            if frame_col:
                break
        if not frame_col:
            for c in df_results.columns:
                if c.lower() == 'label':
                    frame_col = c
                    break
                    
        rebar_col = None
        for c in df_results.columns:
            c_low = c.lower()
            if 'rebar %' in c_low or 'rebar percent' in c_low or 'pmm ratio or rebar' in c_low or 'ratio' in c_low:
                rebar_col = c
                break

        combo_col = None
        for c in df_results.columns:
            c_low = c.lower()
            if 'combo' in c_low:
                combo_col = c
                break

        if frame_col and rebar_col:
            for _, row in df_results.iterrows():
                frame_name = str(row[frame_col]).strip()
                rebar_str = str(row[rebar_col]).strip()
                combo_name = str(row[combo_col]).strip() if combo_col else ""
                rebar_val = 0.0
                try:
                    val_str = rebar_str.replace('%', '').strip()
                    rebar_val = float(val_str)
                    if 'ratio' in rebar_col.lower() and 'rebar %' not in rebar_col.lower() and rebar_val < 0.1 and rebar_val > 0.0:
                        rebar_val *= 100.0
                except ValueError:
                    rebar_val = 0.0
                existing = rebar_data.get(frame_name)
                if existing is None or rebar_val > existing['pct']:
                    rebar_data[frame_name] = {'pct': rebar_val, 'combo': combo_name}
    return rebar_data


def generate_pdf_report(x_grids, y_grids, aggregated_grids, file_path, scale_factor=1.5, low_threshold=2.7, high_threshold=3.0, model_name="Untitled"):
    """Compile final professional landscape engineering PDF reports."""
    doc = SimpleDocTemplate(
        file_path, 
        pagesize=landscape(A3),
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CoverTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=40,
        leading=48,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=20
    )
    
    subtitle_style = ParagraphStyle(
        'CoverSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=15,
        leading=20,
        textColor=colors.HexColor("#475569"),
        spaceAfter=40
    )
    
    h1_style = ParagraphStyle(
        'SectionH1',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=28,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=10,
        keepWithNext=True
    )
    
    body_style = ParagraphStyle(
        'ReportBody',
        parent=styles['BodyText'],
        fontName='Helvetica',
        fontSize=11,
        leading=16,
        textColor=colors.HexColor("#334155"),
        spaceAfter=20
    )
    
    table_header_style = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        textColor=colors.white,
        alignment=1
    )
    
    table_body_style = ParagraphStyle(
        'TableBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        textColor=colors.HexColor("#334155"),
        alignment=1
    )
    
    table_body_bold_style = ParagraphStyle(
        'TableBodyBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        textColor=colors.HexColor("#1E293B"),
        alignment=1
    )
    
    story = []
    
    # ------------------ PAGE 1: COVER ------------------
    story.append(Spacer(1, 100))
    story.append(Paragraph("COLUMN REBAR ANALYSIS REPORT", title_style))
    story.append(Paragraph("Programmatic Extraction & Visualization of Longitudinal Reinforcement from ETABS Models", subtitle_style))
    
    accent_table = Table([['']], colWidths=[1082], rowHeights=[6])
    accent_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#1E3A8A")),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(accent_table)
    story.append(Spacer(1, 60))
    
    meta_title_style = ParagraphStyle('MetaTitle', fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor("#475569"))
    meta_val_style = ParagraphStyle('MetaVal', fontName='Helvetica', fontSize=12, textColor=colors.HexColor("#1E293B"))
    
    metadata = [
        [Paragraph("Project Name:", meta_title_style), Paragraph("Reinforcement Column Audit", meta_val_style)],
        [Paragraph("Source Model:", meta_title_style), Paragraph(model_name, meta_val_style)],
        [Paragraph("Analysis Software:", meta_title_style), Paragraph("CSI ETABS (Active API Interface)", meta_val_style)],
        [Paragraph("Date Generated:", meta_title_style), Paragraph(datetime.now().strftime("%B %d, %Y"), meta_val_style)],
        [Paragraph("Status:", meta_title_style), Paragraph("<font color='#10B981'><b>COMPLETED</b></font>", meta_val_style)]
    ]
    
    meta_table = Table(metadata, colWidths=[180, 902], rowHeights=[26]*5)
    meta_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LINEBELOW', (0,0), (-1,-1), 0.5, colors.HexColor("#F1F5F9")),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 80))
    
    exec_summary_text = (
        "<b>Executive Summary:</b> This report presents the design results of columns extracted directly "
        "from the active ETABS structural analysis model. It maps individual column frame elements across various stories "
        "to their closest horizontal grid points using ETABS bay and connectivity definition tables, and computes the <b>maximum "
        "longitudinal rebar percentage</b> required at each grid location from the design PMM Envelope. The layout drawings on the subsequent "
        "page display the physical shape (rectangular/circular), size, and orientation (local coordinate angle) of the columns. "
        "The colors indicate critical reinforcement demands, helping engineers instantly inspect and audit reinforcement concentrations in the model."
    )
    story.append(Paragraph(exec_summary_text, body_style))
    story.append(PageBreak())
    
    # ------------------ PAGE 2: KEY PLAN ------------------
    story.append(Paragraph("Column Key Plan & Rebar Distribution", h1_style))
    story.append(Paragraph(f"The drawing below shows column geometries in plan rotated based on local axes, alongside color-coded badges indicating maximum rebar percentage. Dynamic threshold ranges are configured as: Green (< {low_threshold:.1f}%), Orange ({low_threshold:.1f}% to {high_threshold:.1f}%), and Red (> {high_threshold:.1f}%).", body_style))
    story.append(Spacer(1, 10))
    
    story.append(ColumnPlanFlowable(x_grids, y_grids, aggregated_grids, width=1082, height=600, scale_factor=scale_factor, low_threshold=low_threshold, high_threshold=high_threshold))
    story.append(PageBreak())
    
    # ------------------ PAGE 3: TABLE ------------------
    story.append(Paragraph("Detailed Reinforcement Data Table", h1_style))
    story.append(Paragraph("Summary of all column coordinates, section details, orientations, and the maximum rebar percentage with governing structural locations.", body_style))
    story.append(Spacer(1, 10))
    
    headers = [
        Paragraph("Grid Point", table_header_style),
        Paragraph("X, Y (m)", table_header_style),
        Paragraph("Drawn Section", table_header_style),
        Paragraph("Shape", table_header_style),
        Paragraph("Dimensions (mm)", table_header_style),
        Paragraph("Angle (°)", table_header_style),
        Paragraph("Max Rebar %", table_header_style),
        Paragraph("PMM Combo", table_header_style),
        Paragraph("Governing Story", table_header_style),
        Paragraph("Frame Element", table_header_style)
    ]
    
    table_rows = [headers]
    for col in aggregated_grids:
        rebar = col['max_rebar']
        
        if rebar == 0.0:
            rebar_p = Paragraph("<font color='#64748B'><b>N/D</b></font>", table_body_bold_style)
        elif rebar < low_threshold:
            rebar_p = Paragraph(f"<font color='#10B981'><b>{rebar:.2f}%</b></font>", table_body_bold_style)
        elif rebar <= high_threshold:
            rebar_p = Paragraph(f"<font color='#F59E0B'><b>{rebar:.2f}%</b></font>", table_body_bold_style)
        else:
            rebar_p = Paragraph(f"<font color='#EF4444'><b>{rebar:.2f}%</b></font>", table_body_bold_style)
            
        dim_str = f"{col['width']*1000:.0f} x {col['depth']*1000:.0f}" if col['shape'] != "Circular" else f"Dia {col['width']*1000:.0f}"
        pmm_combo = col.get('max_rebar_combo', '') or "N/A"

        row_data = [
            Paragraph(col['grid_label'], table_body_bold_style),
            Paragraph(f"{col['x']:.1f}, {col['y']:.1f}", table_body_style),
            Paragraph(col['prop_name'], table_body_style),
            Paragraph(col['shape'], table_body_style),
            Paragraph(dim_str, table_body_style),
            Paragraph(f"{col['angle']:.1f}", table_body_style),
            rebar_p,
            Paragraph(pmm_combo, table_body_style),
            Paragraph(col['max_rebar_story'], table_body_style),
            Paragraph(col['max_rebar_frame'], table_body_style)
        ]
        table_rows.append(row_data)

    col_widths = [100, 140, 130, 70, 130, 70, 90, 120, 130, 102]
    data_table = Table(table_rows, colWidths=col_widths, repeatRows=1)
    
    t_style = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1E3A8A")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
    ]
    
    for idx in range(1, len(table_rows)):
        bg = colors.HexColor("#F8FAFC") if idx % 2 == 0 else colors.HexColor("#FFFFFF")
        t_style.append(('BACKGROUND', (0, idx), (-1, idx), bg))
        
    data_table.setStyle(TableStyle(t_style))
    story.append(data_table)
    
    NumberedCanvas.model_name = model_name
    doc.build(story, canvasmaker=NumberedCanvas)


class ConsoleRedirector(io.StringIO):
    """Thread-safe redirector to update customtkinter textbox."""
    def __init__(self, text_widget, original_stream=None):
        super().__init__()
        self.text_widget = text_widget
        self.original_stream = original_stream
        
    def write(self, s):
        self.text_widget.configure(state="normal")
        self.text_widget.insert("end", s)
        self.text_widget.see("end")
        self.text_widget.configure(state="disabled")
        if self.original_stream is not None:
            try:
                self.original_stream.write(s)
            except Exception:
                pass
        
    def flush(self):
        if self.original_stream is not None:
            try:
                self.original_stream.flush()
            except Exception:
                pass


class EtabsRebarGui(ctk.CTk):
    """CustomTkinter application dashboard interface."""
    def __init__(self):
        super().__init__()
        self.title("ETABS Column Rebar Auditor Dashboard")
        self.geometry("900x650")
        self.resizable(False, False)
        
        self.active_instances = []
        self.selected_moniker = None
        self.out_pdf_filepath = os.path.abspath("Column_Rebar_Report.pdf")
        
        # Grid Configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        self.create_widgets()
        self.scan_etabs_instances()
        
        # Setup thread-safe Queue for GUI updates from console redirection
        sys.stdout = ConsoleRedirector(self.console_box, sys.__stdout__)
        sys.stderr = ConsoleRedirector(self.console_box, sys.__stderr__)
        
    def create_widgets(self):
        # 1. Header Title Frame
        self.header_frame = ctk.CTkFrame(self, height=75, corner_radius=0, fg_color="#1E293B")
        self.header_frame.grid(row=0, column=0, sticky="nsew")
        
        self.title_label = ctk.CTkLabel(self.header_frame, text="ETABS COLUMN REBAR ANALYSIS DASHBOARD", font=ctk.CTkFont(size=20, weight="bold"), text_color="#F8FAFC")
        self.title_label.pack(pady=(12, 2))
        self.subtitle_label = ctk.CTkLabel(self.header_frame, text="Automatic Reinforcement Extraction, Color-Coding & Premium A3 Vector PDF Layout", font=ctk.CTkFont(size=12), text_color="#94A3B8")
        self.subtitle_label.pack()
        
        # 2. Main Dashboard Workspace
        self.main_workspace = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_workspace.grid(row=1, column=0, sticky="nsew", padx=15, pady=15)
        
        self.main_workspace.grid_columnconfigure(0, weight=4) # Left Column: Configuration Controls
        self.main_workspace.grid_columnconfigure(1, weight=5) # Right Column: Terminal Logging Console
        self.main_workspace.grid_rowconfigure(0, weight=1)
        
        # LEFT COLUMN FRAME: CONTROLS
        self.controls_frame = ctk.CTkScrollableFrame(self.main_workspace, fg_color="#0F172A", corner_radius=8)
        self.controls_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Active Model Selector Title
        self.etabs_label = ctk.CTkLabel(self.controls_frame, text="1. SELECT ACTIVE ETABS MODEL", font=ctk.CTkFont(size=14, weight="bold"), text_color="#60A5FA")
        self.etabs_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.selector_inner = ctk.CTkFrame(self.controls_frame, fg_color="transparent")
        self.selector_inner.pack(fill="x", padx=10)
        
        self.instance_combo = ctk.CTkComboBox(self.selector_inner, values=["No instances detected"], width=230, command=self.on_select_instance)
        self.instance_combo.pack(side="left", fill="x", expand=True)
        
        self.refresh_btn = ctk.CTkButton(self.selector_inner, text="↻ Scan", width=65, fg_color="#1E3A8A", hover_color="#2563EB", command=self.scan_etabs_instances)
        self.refresh_btn.pack(side="right", padx=(10, 0))
        
        # Sizing / Layout Configurations
        self.layout_label = ctk.CTkLabel(self.controls_frame, text="2. DRAWING CONFIGURATIONS", font=ctk.CTkFont(size=14, weight="bold"), text_color="#60A5FA")
        self.layout_label.pack(anchor="w", padx=10, pady=(20, 5))
        
        # Scale Factor slider frame
        self.scale_frame = ctk.CTkFrame(self.controls_frame, fg_color="transparent")
        self.scale_frame.pack(fill="x", padx=10, pady=5)
        self.scale_lbl = ctk.CTkLabel(self.scale_frame, text="Visual Scale Factor: 1.5x")
        self.scale_lbl.pack(anchor="w")
        self.scale_slider = ctk.CTkSlider(self.scale_frame, from_=0.5, to=3.0, number_of_steps=25, command=self.on_scale_slider)
        self.scale_slider.set(1.5)
        self.scale_slider.pack(fill="x", pady=2)
        
        # Output Directory Field
        self.save_label = ctk.CTkLabel(self.controls_frame, text="3. PDF SAVE DIRECTORY", font=ctk.CTkFont(size=14, weight="bold"), text_color="#60A5FA")
        self.save_label.pack(anchor="w", padx=10, pady=(20, 5))
        
        self.save_inner = ctk.CTkFrame(self.controls_frame, fg_color="transparent")
        self.save_inner.pack(fill="x", padx=10, pady=(0, 20))
        self.save_path_entry = ctk.CTkEntry(self.save_inner, placeholder_text="Report saving location...")
        self.save_path_entry.insert(0, self.out_pdf_filepath)
        self.save_path_entry.pack(side="left", fill="x", expand=True)
        self.browse_btn = ctk.CTkButton(self.save_inner, text="Browse", width=65, fg_color="#334155", hover_color="#475569", command=self.browse_filepath)
        self.browse_btn.pack(side="right", padx=(10, 0))
        
        # RIGHT COLUMN FRAME: TERMINAL LOGGING CONSOLE
        self.console_frame = ctk.CTkFrame(self.main_workspace, fg_color="#0F172A", corner_radius=8)
        self.console_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        
        self.console_title = ctk.CTkLabel(self.console_frame, text="EXECUTION SYSTEM LOGS", font=ctk.CTkFont(size=13, weight="bold"), text_color="#94A3B8")
        self.console_title.pack(anchor="w", padx=15, pady=(10, 5))
        
        self.console_box = ctk.CTkTextbox(self.console_frame, font=ctk.CTkFont(family="Consolas", size=10), text_color="#10B981", fg_color="#020617", state="disabled")
        self.console_box.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
        # 3. Action Footer Frame
        self.footer_frame = ctk.CTkFrame(self, height=80, corner_radius=0, fg_color="#1E293B")
        self.footer_frame.grid(row=2, column=0, sticky="nsew")
        
        self.progress_bar = ctk.CTkProgressBar(self.footer_frame, width=870, fg_color="#334155", progress_color="#2563EB", height=8)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=(10, 10), padx=15)
        
        self.footer_actions = ctk.CTkFrame(self.footer_frame, fg_color="transparent")
        self.footer_actions.pack(fill="x", padx=15, pady=(0, 15))
        
        self.run_btn = ctk.CTkButton(self.footer_actions, text="🚀 Generate Premium Rebar Report", width=350, font=ctk.CTkFont(size=14, weight="bold"), fg_color="#2563EB", hover_color="#1D4ED8", command=self.start_generation_process)
        self.run_btn.pack(side="left")
        
        self.open_pdf_btn = ctk.CTkButton(self.footer_actions, text="📂 Open Report PDF", width=160, font=ctk.CTkFont(size=14, weight="bold"), fg_color="#059669", hover_color="#047857", state="disabled", command=self.open_generated_pdf)
        self.open_pdf_btn.pack(side="right", padx=(10, 0))
        
        self.open_folder_btn = ctk.CTkButton(self.footer_actions, text="📁 Open Folder", width=140, font=ctk.CTkFont(size=14, weight="bold"), fg_color="#4B5563", hover_color="#374151", state="disabled", command=self.open_output_folder)
        self.open_folder_btn.pack(side="right")
        
    def scan_etabs_instances(self):
        print("[GUI] Scanning active Running Object Table (ROT)...")
        try:
            rot = pythoncom.GetRunningObjectTable()
            bind_ctx = pythoncom.CreateBindCtx(0)
            enum_monikers = rot.EnumRunning()
        except Exception as e:
            messagebox.showerror("System Error", f"Failed to access ROT registry table: {e}")
            return
            
        self.active_instances = []
        combo_vals = []
        
        for moniker in enum_monikers:
            try:
                name = moniker.GetDisplayName(bind_ctx, None)
                if "csi.etabs.api.etabsobject" in name.lower():
                    pid = name.split(":")[-1]
                    punk_pywin = rot.GetObject(moniker)
                    
                    rep = repr(punk_pywin)
                    match = re.search(r"with obj at (0x[0-9a-fA-F]+)", rep)
                    if not match:
                        continue
                    addr_val = int(match.group(1), 16)
                    
                    # Bind via comtypes to fetch name
                    punk = ctypes.cast(addr_val, POINTER(IUnknown))
                    punk.AddRef()
                    etabs_object = comtypes.client.wrap(punk)
                    sap_model = etabs_object.SapModel
                    
                    try:
                        full_filepath = sap_model.GetModelFilename(True)
                        filename = os.path.basename(full_filepath) if full_filepath else "Untitled"
                    except Exception:
                        full_filepath = ""
                        filename = "Untitled / No Model Loaded"
                        
                    self.active_instances.append({
                        'pid': pid,
                        'filename': filename,
                        'full_filepath': full_filepath,
                        'addr': addr_val,
                        'moniker_name': name
                    })
                    combo_vals.append(filename)
            except Exception:
                continue
                
        if self.active_instances:
            self.instance_combo.configure(values=combo_vals)
            self.instance_combo.set(combo_vals[0])
            self.selected_moniker = self.active_instances[0]
            print(f"[GUI] Resolved {len(self.active_instances)} active ETABS instances successfully.")
            self.update_default_save_path()
        else:
            self.instance_combo.configure(values=["No instances detected"])
            self.instance_combo.set("No instances detected")
            self.selected_moniker = None
            print("[GUI] No active ETABS models found in the system registry.")
            
    def on_select_instance(self, val_str):
        if not self.active_instances:
            return
        for inst in self.active_instances:
            if inst['filename'] == val_str:
                self.selected_moniker = inst
                print(f"[GUI] Selected target instance: PID {inst['pid']} - {inst['filename']}")
                self.update_default_save_path()
                break
                
    def update_default_save_path(self):
        if self.selected_moniker and self.selected_moniker['full_filepath']:
            model_dir = os.path.dirname(self.selected_moniker['full_filepath'])
            if os.path.isdir(model_dir):
                self.out_pdf_filepath = os.path.abspath(os.path.join(model_dir, "Column_Rebar_Report.pdf"))
                if hasattr(self, 'save_path_entry'):
                    self.save_path_entry.delete(0, "end")
                    self.save_path_entry.insert(0, self.out_pdf_filepath)
                
    def on_scale_slider(self, val):
        self.scale_lbl.configure(text=f"Visual Scale Factor: {val:.2f}x")
        
    def browse_filepath(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")], initialfile=os.path.basename(self.out_pdf_filepath))
        if path:
            self.out_pdf_filepath = os.path.abspath(path)
            self.save_path_entry.delete(0, "end")
            self.save_path_entry.insert(0, self.out_pdf_filepath)
            
    def start_generation_process(self):
        if not self.selected_moniker:
            messagebox.showwarning("Active Link Required", "Please ensure ETABS is running and an active model instance is selected before executing the reinforcement analysis.")
            return
            
        self.run_btn.configure(state="disabled", text="⚡ Processing Model Data...")
        self.open_pdf_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.progress_bar.set(0.05)
        
        # Read sliders/fields
        scale_fac = self.scale_slider.get()
        low_th = 2.7
        high_th = 3.0
        out_path = self.save_path_entry.get().strip()
        
        if not out_path:
            out_path = self.out_pdf_filepath
            
        # Launch analysis in background thread to avoid freezing GUI
        t = threading.Thread(target=self.run_background_analyzer, args=(scale_fac, low_th, high_th, out_path))
        t.daemon = True
        t.start()
        
    def run_background_analyzer(self, scale_fac, low_th, high_th, out_path):
        print("="*60)
        print("     STARTING ANALYSIS & GRAPHICAL LAYOUT REPORT GENERATION     ")
        print("="*60)
        print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Target PID: {self.selected_moniker['pid']}")
        print(f"Target Model: {self.selected_moniker['filename']}")
        print(f"Output File: {out_path}")
        print(f"Configurations: Scale={scale_fac:.2f}x | Green < {low_th:.2f}% | Red > {high_th:.2f}%\n")
        
        sap_model = None
        original_units = None
        
        try:
            self.progress_bar.set(0.15)
            # Step 1: Bind OAPI to the selected instance address
            print("Connecting to ETABS API...")
            pythoncom.CoInitialize()
            
            # Retrieve the COM object directly in this background thread by PID
            target_pid = self.selected_moniker['pid']
            rot = pythoncom.GetRunningObjectTable()
            bind_ctx = pythoncom.CreateBindCtx(0)
            enum_monikers = rot.EnumRunning()
            
            punk_pywin = None
            for moniker in enum_monikers:
                try:
                    name = moniker.GetDisplayName(bind_ctx, None)
                    if "csi.etabs.api.etabsobject" in name.lower():
                        pid = name.split(":")[-1]
                        if pid == str(target_pid):
                            punk_pywin = rot.GetObject(moniker)
                            break
                except Exception:
                    continue
            
            if not punk_pywin:
                raise RuntimeError(f"Could not connect to ETABS instance with PID {target_pid} from background thread.")
                
            rep = repr(punk_pywin)
            match = re.search(r"with obj at (0x[0-9a-fA-F]+)", rep)
            if not match:
                raise RuntimeError("Failed to extract COM interface address from PyIDispatch.")
                
            addr_val = int(match.group(1), 16)
            punk = ctypes.cast(addr_val, POINTER(IUnknown))
            punk.AddRef()
            etabs_object = comtypes.client.wrap(punk)
            sap_model = etabs_object.SapModel
            print("Connected to active ETABS COM framework successfully.")
            
            self.progress_bar.set(0.25)
            # Step 2: Configure metrics units
            original_units = sap_model.GetPresentUnits()
            sap_model.SetPresentUnits(6) # 6 = kN_m_C
            print("Present units set to: Kilonewtons, Meters, Celsius.")
            
            self.progress_bar.set(0.35)
            # Step 3: Read required coordinate and mapping tables
            print("Extracting model structural connectivity...")
            df_connectivity = read_etabs_table(sap_model, "Column Object Connectivity")
            df_column_bays = read_etabs_table(sap_model, "Column Bays")
            df_point_bays = read_etabs_table(sap_model, "Point Bays")
            df_grids = read_etabs_table(sap_model, "Grid Definitions - Grid Lines")
            
            if any(df is None for df in [df_connectivity, df_column_bays, df_point_bays, df_grids]):
                raise RuntimeError("Failed to read column coordinate mapping tables from database.")
                
            df_connectivity.columns = [c.strip() for c in df_connectivity.columns]
            df_column_bays.columns = [c.strip() for c in df_column_bays.columns]
            df_point_bays.columns = [c.strip() for c in df_point_bays.columns]
            df_grids.columns = [c.strip() for c in df_grids.columns]
            
            self.progress_bar.set(0.45)
            # Step 4: Scan structural Grid lines
            x_grids, y_grids = get_grid_lines(sap_model)
            if not x_grids or not y_grids:
                raise RuntimeError("Failed to resolve coordinate grid limits.")
                
            self.progress_bar.set(0.55)
            # Step 5: Query rebar ratios from concrete summary database envelope
            print("Extracting reinforcement ratios from design PMM database Envelope...")
            rebar_data = extract_rebar_design_results(sap_model)
            print(f"Extracted design results for {len(rebar_data)} frames.")
            
            self.progress_bar.set(0.65)
            # Step 6: Map column frames and orientations
            print("Processing columns coordinates and geometries...")
            col_bay_to_pt_bay = {str(row['Label']).strip(): str(row['PointBayI']).strip() for _, row in df_column_bays.iterrows()}
            pt_bay_to_coord = {str(row['Label']).strip(): (float(row['X']), float(row['Y'])) for _, row in df_point_bays.iterrows()}
            
            resolved_cols = []
            for _, row in df_connectivity.iterrows():
                uniq_name = str(row['Unique Name']).strip()
                story = str(row['Story']).strip()
                col_bay = str(row['ColumnBay']).strip()
                
                pt_bay = col_bay_to_pt_bay.get(col_bay, None)
                if not pt_bay:
                    continue
                coord = pt_bay_to_coord.get(pt_bay, None)
                if not coord:
                    continue
                    
                cx, cy = coord
                
                # Match closest grid intersection
                best_x_id = "Off-Grid"
                min_dist_x = float('inf')
                for gid, ox in x_grids:
                    dist = abs(cx - ox)
                    if dist < min_dist_x and dist < 0.5:
                        min_dist_x = dist
                        best_x_id = gid
                        
                best_y_id = "Off-Grid"
                min_dist_y = float('inf')
                for gid, oy in y_grids:
                    dist = abs(cy - oy)
                    if dist < min_dist_y and dist < 0.5:
                        min_dist_y = dist
                        best_y_id = gid
                        
                grid_label = f"{best_x_id}-{best_y_id}"
                
                try:
                    ret_sect = sap_model.FrameObj.GetSection(uniq_name)
                    section_name = ret_sect[0]
                except Exception:
                    section_name = "Default"
                    
                shape, w, d = get_section_dimensions(sap_model, section_name)
                
                angle = 0.0
                try:
                    ret_axes = sap_model.FrameObj.GetLocalAxes(uniq_name)
                    if ret_axes[-1] == 0:
                        angle = ret_axes[0]
                except Exception:
                    pass
                    
                rebar_info = rebar_data.get(uniq_name, {})
                rebar_pct = rebar_info.get('pct', 0.0)
                pmm_combo = rebar_info.get('combo', '')

                resolved_cols.append({
                    'frame_name': uniq_name,
                    'col_label': col_bay,
                    'story_name': story,
                    'x': cx,
                    'y': cy,
                    'grid_label': grid_label,
                    'prop_name': section_name,
                    'shape': shape,
                    'width': w,
                    'depth': d,
                    'angle': angle,
                    'rebar_pct': rebar_pct,
                    'pmm_combo': pmm_combo
                })
                
            if not resolved_cols:
                raise RuntimeError("No column frame elements resolved from database structure.")
                
            # Perform Layer 3 individual COM summary query fallback if required
            zero_rebar_cols = [c for c in resolved_cols if c['rebar_pct'] == 0.0]
            if 0 < len(zero_rebar_cols) < len(resolved_cols):
                print(f"Layer 3 fallback: querying {len(zero_rebar_cols)} columns individually...")
                for col in zero_rebar_cols:
                    try:
                        ret = sap_model.DesignConcrete.GetSummaryResultsColumn(col['frame_name'])
                        if isinstance(ret, (tuple, list)) and len(ret) > 13:
                            num_res = ret[0]
                            top_area = ret[3]
                            bot_area = ret[4]
                            ret_code = ret[-1]
                            
                            if ret_code == 0 and num_res > 0:
                                top_f = [float(x) for x in top_area if str(x).replace('.','',1).isdigit()]
                                bot_f = [float(x) for x in bot_area if str(x).replace('.','',1).isdigit()]
                                combined = top_f + bot_f
                                if combined:
                                    max_a = max(combined)
                                    w, d = col['width'], col['depth']
                                    sect_area = np.pi * (w/2.0)**2 if col['shape'] == "Circular" else w * d
                                    pct = (max_a / sect_area) * 100.0
                                    if pct > 10.0 and sect_area < 10.0 and max_a > 1.0:
                                        pct = (max_a / (sect_area * 1e6)) * 100.0
                                    col['rebar_pct'] = pct
                    except Exception:
                        pass
                        
            self.progress_bar.set(0.75)
            # Step 7: Group columns by horizontal grid intersection coordinates
            grid_groups = {}
            for col in resolved_cols:
                key = (col['grid_label'], col['x'], col['y'])
                if key not in grid_groups:
                    grid_groups[key] = []
                grid_groups[key].append(col)
                
            aggregated_grids = []
            for (grid_label, gx, gy), cols in grid_groups.items():
                max_rebar_col = max(cols, key=lambda c: c['rebar_pct'])
                
                def get_area(c):
                    return np.pi * (c['width']/2.0)**2 if c['shape'] == "Circular" else c['width'] * c['depth']
                largest_col = max(cols, key=get_area)
                
                aggregated_grids.append({
                    'grid_label': grid_label,
                    'x': gx,
                    'y': gy,
                    'max_rebar': max_rebar_col['rebar_pct'],
                    'max_rebar_story': max_rebar_col['story_name'],
                    'max_rebar_frame': max_rebar_col['frame_name'],
                    'max_rebar_combo': max_rebar_col.get('pmm_combo', ''),
                    'shape': largest_col['shape'],
                    'width': largest_col['width'],
                    'depth': largest_col['depth'],
                    'angle': largest_col['angle'],
                    'prop_name': largest_col['prop_name']
                })
            aggregated_grids = sorted(aggregated_grids, key=lambda x: x['grid_label'])
            
            self.progress_bar.set(0.85)
            # Step 8: Build the PDF Report with robust file-lock fallback
            print("Compiling landscape vector drawing layout sheets...")
            
            pdf_file = out_path
            attempt = 1
            while True:
                try:
                    generate_pdf_report(
                        x_grids, y_grids, aggregated_grids, pdf_file,
                        scale_factor=scale_fac, low_threshold=low_th, high_threshold=high_th,
                        model_name=self.selected_moniker['filename']
                    )
                    break
                except PermissionError:
                    base, ext = os.path.splitext(out_path)
                    pdf_file = f"{base}_{attempt}{ext}"
                    print(f"Warning: Destination file locked. Retrying as '{os.path.basename(pdf_file)}'...")
                    attempt += 1
                    
            self.out_pdf_filepath = os.path.abspath(pdf_file)
            
            self.progress_bar.set(1.0)
            # Restore ETABS units
            sap_model.SetPresentUnits(original_units)
            print("\nETABS COM units successfully restored to original state.")
            print("="*60)
            print("SUCCESS! COLUMN REBAR SUMMARY SHEET CREATION COMPLETED.")
            print(f"Report Output: {self.out_pdf_filepath}")
            print("="*60)
            
            # Update GUI states back on main thread
            self.after(0, self.on_generation_success)
            
        except Exception as err:
            err_msg = str(err)
            print(f"\n[FATAL ERROR] Analysis crashed: {err_msg}")
            print("="*60)
            if sap_model and original_units:
                try:
                    sap_model.SetPresentUnits(original_units)
                except Exception:
                    pass
            self.after(0, lambda: self.on_generation_failure(err_msg))
            
    def on_generation_success(self):
        self.run_btn.configure(state="normal", text="🚀 Generate Premium Rebar Report")
        self.open_pdf_btn.configure(state="normal")
        self.open_folder_btn.configure(state="normal")
        messagebox.showinfo("Analysis Finished", f"The reinforcement analysis was completed successfully!\n\nPDF Location:\n{self.out_pdf_filepath}")
        
    def on_generation_failure(self, err):
        self.run_btn.configure(state="normal", text="🚀 Generate Premium Rebar Report")
        self.progress_bar.set(0.0)
        messagebox.showerror("Execution Crash", f"An exception occurred during ETABS data extraction:\n\n{err}")
        
    def open_generated_pdf(self):
        if os.path.exists(self.out_pdf_filepath):
            print(f"[GUI] Opening PDF report file in system default viewer: {os.path.basename(self.out_pdf_filepath)}")
            os.startfile(self.out_pdf_filepath)
        else:
            messagebox.showerror("File Error", "The PDF report file could not be found on disk.")
            
    def open_output_folder(self):
        if self.out_pdf_filepath:
            pdf_dir = os.path.dirname(self.out_pdf_filepath)
            if os.path.isdir(pdf_dir):
                print(f"[GUI] Opening output directory in Explorer: {pdf_dir}")
                os.startfile(pdf_dir)
            else:
                messagebox.showerror("Folder Error", "The output folder could not be found on disk.")
        else:
            messagebox.showerror("Folder Error", "No output path configured.")


if __name__ == "__main__":
    app = EtabsRebarGui()
    app.mainloop()
