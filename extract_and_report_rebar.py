import comtypes.client
import pandas as pd
import numpy as np
import io
import os
import sys
from datetime import datetime

# ReportLab imports for generating a premium PDF
from reportlab.lib.pagesizes import letter, A4, A3, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
from reportlab.pdfgen import canvas
from reportlab.platypus import Flowable

class NumberedCanvas(canvas.Canvas):
    """
    Custom canvas to enable 2-pass page numbering and professional header/footers 
    on A3 landscape pages.
    """
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
        self.drawString(54, 30, "Confidential - Engineering Design Document")
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
    def __init__(self, x_grids, y_grids, aggregated_grids, width, height):
        super().__init__()
        self.x_grids = x_grids
        self.y_grids = y_grids
        self.aggregated_grids = aggregated_grids
        self.width = width
        self.height = height
        
    def wrap(self, availWidth, availHeight):
        return self.width, self.height
        
    def draw(self):
        # In ReportLab, the active canvas is accessed via self.canv
        canvas = self.canv
        
        # Gather all X and Y coordinates to determine bounding box
        x_coords = [val for _, val in self.x_grids]
        y_coords = [val for _, val in self.y_grids]
        
        for col in self.aggregated_grids:
            x_coords.append(col['x'])
            y_coords.append(col['y'])
            
        if not x_coords or not y_coords:
            # Nothing to draw
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
        bubble_r = 12
        canvas.setFont("Helvetica-Bold", 11)
        
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
            canvas.drawCentredString(x_pdf, y_start_pdf - bubble_r - 4.0, label)
            
            # Solid bubble at top
            canvas.setFillColor(colors.HexColor("#E2E8F0"))
            canvas.circle(x_pdf, y_end_pdf + bubble_r, bubble_r, stroke=1, fill=1)
            canvas.setFillColor(colors.HexColor("#1E293B"))
            canvas.drawCentredString(x_pdf, y_end_pdf + bubble_r - 4.0, label)
            
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
            canvas.drawCentredString(x_start_pdf - bubble_r, y_pdf - 4.0, label)
            
            # Solid bubble at right
            canvas.setFillColor(colors.HexColor("#E2E8F0"))
            canvas.circle(x_end_pdf + bubble_r, y_pdf, bubble_r, stroke=1, fill=1)
            canvas.setFillColor(colors.HexColor("#1E293B"))
            canvas.drawCentredString(x_end_pdf + bubble_r, y_pdf - 4.0, label)
            
        # Refined column scaling factor: scale down to prevent geometry overlap
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
                if dist > 0.05 and dist < min_dist_m: # Ignore duplicate or extremely close modeling errors (< 5cm)
                    min_dist_m = dist
                    
        # Determine the maximum allowed column drawing size based on spacing and limits
        max_col_size_pt = 6.0  # Sleek default maximum size
        if min_dist_m != float('inf'):
            min_dist_pt = min_dist_m * scale
            # To prevent geometric overlap, columns should be separated by a gap
            max_col_size_pt = min(6.0, max(2.7, min_dist_pt * 0.7 * 1.5))
            
        # Determine the minimum allowed column drawing size
        min_col_size_pt = min(2.7, max_col_size_pt)
        
        # Enforce column drawing size within bounds
        if avg_col_size_pt < min_col_size_pt:
            col_scale = scale * (min_col_size_pt / avg_col_size_pt)
            
        if col_scale * avg_col_size_m > max_col_size_pt:
            col_scale = max_col_size_pt / avg_col_size_m
            
        print(f"[DIAGNOSTIC] scale={scale:.4f}, min_dist_m={min_dist_m:.4f}, max_col_size_pt={max_col_size_pt:.4f}, col_scale={col_scale:.4f}, final_avg_col_size_pt={col_scale * avg_col_size_m:.4f}")
        
        occupied_rects = []
        
        # Bounding box calculation for columns to reserve their spaces first
        for col in self.aggregated_grids:
            cx, cy = to_pdf(col['x'], col['y'])
            w = col['width'] * col_scale
            d = col['depth'] * col_scale
            r = max(w, d) / 2.0
            # Pad the column space slightly
            col_rect = (cx - r - 1.5, cy - r - 1.5, cx + r + 1.5, cy + r + 1.5)
            occupied_rects.append(col_rect)
            
        # 3. Draw Column Shapes
        for col in self.aggregated_grids:
            cx, cy = to_pdf(col['x'], col['y'])
            w = col['width'] * col_scale
            d = col['depth'] * col_scale
            shape = col['shape']
            angle = col['angle']
            
            # Drawing column shape
            canvas.saveState()
            canvas.translate(cx, cy)
            canvas.rotate(angle)
            
            canvas.setFillColor(colors.HexColor("#1E293B")) # Elegant charcoal black column fill
            canvas.setStrokeColor(colors.HexColor("#0F172A"))
            canvas.setLineWidth(0.8)
            canvas.setDash()
            
            if shape == "Circular":
                r = w / 2.0
                canvas.circle(0, 0, r, stroke=1, fill=1)
                # Draw internal crossing lines for engineering details
                canvas.setStrokeColor(colors.HexColor("#475569"))
                canvas.setLineWidth(0.4)
                canvas.line(-r, 0, r, 0)
                canvas.line(0, -r, 0, r)
            else: # Rectangular column
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
            shape = col['shape']
            
            rebar = col['max_rebar']
            
            # Select color scheme based on rebar %
            if rebar == 0.0:
                bg_color = colors.HexColor("#64748B") # Slate grey for no design
                text_str = "N/D"
            elif rebar < 2.7:
                bg_color = colors.HexColor("#10B981") # Green
                text_str = f"{rebar:.2f}%"
            elif rebar <= 3.0:
                bg_color = colors.HexColor("#F59E0B") # Orange
                text_str = f"{rebar:.2f}%"
            else:
                bg_color = colors.HexColor("#EF4444") # Red
                text_str = f"{rebar:.2f}%"
                
            canvas.setFont("Helvetica-Bold", 9.75)
            text_width = canvas.stringWidth(text_str, "Helvetica-Bold", 9.75)
            badge_w = text_width + 6
            badge_h = 13.5
            
            r_col = max(w, d) / 2.0
            
            # Bounded candidate offset directions for labels: TR, BL, TL, BR, T, B, R, L
            candidates = [
                (r_col + 2, r_col + 2), # Top Right
                (-r_col - badge_w - 2, -r_col - badge_h - 2), # Bottom Left
                (-r_col - badge_w - 2, r_col + 2), # Top Left
                (r_col + 2, -r_col - badge_h - 2), # Bottom Right
                (-badge_w/2, r_col + 3), # Top Center
                (-badge_w/2, -r_col - badge_h - 3), # Bottom Center
                (r_col + 3, -badge_h/2), # Right Center
                (-r_col - badge_w - 3, -badge_h/2) # Left Center
            ]
            
            selected_pos = None
            for dx, dy in candidates:
                bx = cx + dx
                by = cy + dy
                candidate_rect = (bx, by, bx + badge_w, by + badge_h)
                
                # Check collision with all occupied bounds (other columns/badges)
                has_collision = False
                for r_occ in occupied_rects:
                    if rects_intersect(candidate_rect, r_occ, margin=0.5):
                        has_collision = True
                        break
                        
                if not has_collision:
                    selected_pos = (bx, by, candidate_rect)
                    break
                    
            if not selected_pos:
                # If everything intersects, fallback to Top-Right but offset slightly further out
                bx = cx + r_col + 4
                by = cy + r_col + 4
                candidate_rect = (bx, by, bx + badge_w, by + badge_h)
                selected_pos = (bx, by, candidate_rect)
                
            bx, by, final_rect = selected_pos
            occupied_rects.append(final_rect)
            
            # Draw badge background
            canvas.setFillColor(bg_color)
            canvas.roundRect(bx, by, badge_w, badge_h, 2, stroke=0, fill=1)
            
            # Print badge text in white
            canvas.setFillColor(colors.white)
            canvas.drawString(bx + 3, by + 3.0, text_str)
            
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
            
            # Print Column Section Size label centered below the column (e.g. 40X50 for 400X500)
            canvas.setFillColor(colors.HexColor("#475569"))
            canvas.setFont("Helvetica-Bold", 9.0)
            lbl_offset = r_col + 9
            if col['shape'] == "Circular":
                size_str = f"Dia {col['width']*100:.0f}"
            else:
                size_str = f"{col['width']*100:.0f}X{col['depth']*100:.0f}"
            canvas.drawCentredString(cx, cy - lbl_offset, size_str)
            
        # 5. Draw Legend Box at bottom right of the canvas
        legend_x = self.width - 215
        legend_y = 15
        legend_w = 200
        legend_h = 95
        
        canvas.setStrokeColor(colors.HexColor("#CBD5E1"))
        canvas.setLineWidth(0.5)
        canvas.setFillColor(colors.HexColor("#FFFFFF"))
        canvas.roundRect(legend_x, legend_y, legend_w, legend_h, 4, stroke=1, fill=1)
        
        canvas.setFont("Helvetica-Bold", 11.0)
        canvas.setFillColor(colors.HexColor("#1E293B"))
        canvas.drawString(legend_x + 10, legend_y + legend_h - 15, "LEGEND: Rebar %")
        
        entries = [
            (colors.HexColor("#10B981"), "Rebar < 2.7% (Low)"),
            (colors.HexColor("#F59E0B"), "2.7% <= Rebar <= 3.0%"),
            (colors.HexColor("#EF4444"), "Rebar > 3.0% (High)"),
            (colors.HexColor("#64748B"), "N/D (Not Designed)")
        ]
        
        for idx, (color, label_text) in enumerate(entries):
            ey = legend_y + legend_h - 33 - idx * 15.5
            # Color indicator badge
            canvas.setFillColor(color)
            canvas.roundRect(legend_x + 10, ey, 30, 10, 2, stroke=0, fill=1)
            # Label
            canvas.setFillColor(colors.HexColor("#475569"))
            canvas.setFont("Helvetica", 9.0)
            canvas.drawString(legend_x + 46, ey + 1.5, label_text)
            
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
        canvas.setFont("Helvetica", 6)
        canvas.setFillColor(colors.HexColor("#475569"))
        canvas.drawString(scale_x, scale_y + 4, f"Scale Bar: {scale_m:.1f} m")


def connect_to_etabs():
    """
    Connect to the active instance of ETABS and return SapModel.
    Scans the ROT using pywin32 to let the user select from multiple running instances.
    If only one instance is found or pywin32 is unavailable, falls back gracefully.
    """
    print("Connecting to ETABS...")
    
    # Try pywin32 ROT-based multi-instance selection first
    try:
        import pythoncom
        import ctypes
        import re
        
        rot = pythoncom.GetRunningObjectTable()
        bind_ctx = pythoncom.CreateBindCtx(0)
        enum_monikers = rot.EnumRunning()
        
        instances = []
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
                    
                    # Bind using comtypes
                    from comtypes import IUnknown, POINTER
                    punk = ctypes.cast(addr_val, POINTER(IUnknown))
                    punk.AddRef()
                    
                    etabs_object = comtypes.client.wrap(punk)
                    sap_model = etabs_object.SapModel
                    
                    try:
                        filename = sap_model.GetModelFilename(False)
                        filepath = sap_model.GetModelFilename(True)
                    except Exception:
                        filename = "Untitled / No Model Loaded"
                        filepath = "Untitled / No Model Loaded"
                        
                    instances.append({
                        'pid': pid,
                        'filename': filename,
                        'filepath': filepath,
                        'sap_model': sap_model
                    })
            except Exception:
                continue
                
        if instances:
            if len(instances) == 1:
                inst = instances[0]
                print(f"Connected to running ETABS instance (PID: {inst['pid']}, Model: {inst['filename']})")
                return inst['sap_model']
            else:
                print(f"\nDetected {len(instances)} running ETABS instances:")
                for idx, inst in enumerate(instances):
                    print(f"[{idx + 1}] PID: {inst['pid']} | Model: {inst['filename']}")
                    print(f"    Path: {inst['filepath']}")
                print("-" * 50)
                
                # Selection prompt
                while True:
                    try:
                        choice_str = input(f"Select ETABS instance (1-{len(instances)}) [default 1]: ").strip()
                        if not choice_str:
                            inst = instances[0]
                            break
                        choice = int(choice_str)
                        if 1 <= choice <= len(instances):
                            inst = instances[choice - 1]
                            break
                        else:
                            print(f"Please enter a number between 1 and {len(instances)}.")
                    except ValueError:
                        print("Invalid input. Please enter a valid number.")
                    except (KeyboardInterrupt, EOFError):
                        print("\nSelection cancelled. Defaulting to instance 1.")
                        inst = instances[0]
                        break
                        
                print(f"\nConnected to selected instance (PID: {inst['pid']}, Model: {inst['filename']})")
                return inst['sap_model']
    except Exception as e:
        print(f"ROT-based connection scan bypassed or pywin32 not available: {e}")
        
    # Standard single-instance fallback
    try:
        helper = comtypes.client.CreateObject('ETABSv1.Helper')
        helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
        etabs_object = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
        sap_model = etabs_object.SapModel
        print("Connected to active ETABS object successfully.")
        return sap_model
    except Exception as e:
        print(f"Error connecting to active ETABS: {e}")
        print("Ensure ETABS is running and a model is open with concrete design completed.")
        return None


def get_grid_lines(sap_model):
    """
    Extract coordinate grid lines from ETABS using database tables.
    Returns sorted list of X grids and Y grids [(label, coordinate), ...]
    """
    print("Extracting grid definitions from database tables...")
    table_key = "Grid Definitions - Grid Lines"
    
    df_results = read_etabs_table(sap_model, table_key)
    
    if df_results is None:
        print("Grid definitions table not found or empty. Using virtual grid fallback.")
        return None, None
        
    try:
        df_results.columns = [c.strip() for c in df_results.columns]
        
        x_grids = []
        y_grids = []
        
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
                
        # Sort and unique
        x_grids = sorted(list(set(x_grids)), key=lambda x: x[1])
        y_grids = sorted(list(set(y_grids)), key=lambda x: x[1])
        
        print(f"Extracted {len(x_grids)} X Grid lines and {len(y_grids)} Y Grid lines from ETABS.")
        return x_grids, y_grids
        
    except Exception as e:
        print(f"Failed to parse grid table: {e}. Utilizing virtual grid fallback.")
        return None, None


def generate_virtual_grids(column_coords, tolerance=0.5):
    """
    Generate coordinates for a virtual grid system based on unique column centers.
    """
    if not column_coords:
        return [("X1", 0.0)], [("Y1", 0.0)]
        
    x_coords = sorted(list(set([c[0] for c in column_coords])))
    y_coords = sorted(list(set([c[1] for c in column_coords])))
    
    # Cluster unique X coordinates
    grouped_x = []
    current = []
    for x in x_coords:
        if not current:
            current.append(x)
        elif abs(x - current[-1]) < tolerance:
            current.append(x)
        else:
            grouped_x.append(np.mean(current))
            current = [x]
    if current:
        grouped_x.append(np.mean(current))
        
    # Cluster unique Y coordinates
    grouped_y = []
    current = []
    for y in y_coords:
        if not current:
            current.append(y)
        elif abs(y - current[-1]) < tolerance:
            current.append(y)
        else:
            grouped_y.append(np.mean(current))
            current = [y]
    if current:
        grouped_y.append(np.mean(current))
        
    # Label grids
    import string
    x_grids = []
    for i, x in enumerate(grouped_x):
        label = ""
        n = i
        while n >= 0:
            label = string.ascii_uppercase[n % 26] + label
            n = n // 26 - 1
        x_grids.append((label, x))
        
    y_grids = []
    for i, y in enumerate(grouped_y):
        y_grids.append((str(i+1), y))
        
    print(f"Generated virtual grid layout: {len(x_grids)} X grids, {len(y_grids)} Y grids.")
    return x_grids, y_grids


def get_section_dimensions(sap_model, prop_name):
    """Retrieve the physical width and depth of a section, identifying its shape."""
    # 1. Try Rectangular
    try:
        ret = sap_model.PropFrame.GetRectangle(prop_name)
        if ret[-1] == 0:
            # Correct indices: ret[3] is width (t2), ret[2] is depth (t3)
            return "Rectangular", ret[3], ret[2]
    except Exception:
        pass
        
    # 2. Try Circular
    try:
        ret = sap_model.PropFrame.GetCircle(prop_name)
        if ret[-1] == 0:
            # Correct indices: ret[2] is diameter (t3)
            return "Circular", ret[2], ret[2]
    except Exception:
        pass
        
    # 3. Fallback: Check general section properties table if available, or use defaults
    return "Rectangular", 0.4, 0.4 # Default to 400x400mm rectangular


def extract_rebar_design_results(sap_model):
    """
    Extract required longitudinal rebar percentage for each column frame directly from 
    the Concrete Column PMM Envelope database table.
    """
    print("Searching design database tables for PMM Envelope reinforcement ratios...")
    
    # Layer 1: Query available table names
    table_keys = []
    try:
        ret = sap_model.DatabaseTables.GetAvailableTables()
        num_tables, keys, names, import_types, ret_code = ret
        if ret_code == 0:
            table_keys = list(keys)
    except Exception as e:
        print(f"Could not retrieve available database tables: {e}")
        
    # Scan keys to find any concrete column PMM envelope tables
    design_table_key = None
    for key in table_keys:
        k_lower = key.lower()
        if "concrete column" in k_lower and "pmm" in k_lower and "envelope" in k_lower:
            design_table_key = key
            break
            
    # Direct fallbacks if we couldn't list tables or none matched
    common_keys = [
        "Concrete Column PMM Envelope - IS 456-2000",
        "Concrete Column PMM Envelope - ACI 318-19",
        "Concrete Column PMM Envelope - ACI 318-14",
        "Concrete Column PMM Envelope - Eurocode 2-2004",
        "Concrete Column PMM Envelope"
    ]
    
    df_results = None
    if design_table_key:
        print(f"Found PMM Envelope table: '{design_table_key}'")
        df_results = read_etabs_table(sap_model, design_table_key)
    else:
        for key in common_keys:
            print(f"Trying direct table query for: '{key}'...")
            df_results = read_etabs_table(sap_model, key)
            if df_results is not None:
                design_table_key = key
                break
                
    rebar_data = {} # Map frame_name -> {'pct': max rebar percent, 'combo': governing PMM combo}
    
    if df_results is not None:
        print("Extracting rebar data from PMM Envelope table...")
        df_results.columns = [c.strip() for c in df_results.columns]
        
        # Prioritize UniqueName or Unique Name for frames
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
                    
        # Find rebar % or ratio column
        rebar_col = None
        for c in df_results.columns:
            c_low = c.lower()
            if 'rebar %' in c_low or 'rebar percent' in c_low or 'pmm ratio or rebar' in c_low or 'ratio' in c_low:
                rebar_col = c
                break

        # Find governing PMM combo column
        combo_col = None
        for c in df_results.columns:
            c_low = c.lower()
            if 'combo' in c_low:
                combo_col = c
                break

        if frame_col and rebar_col:
            print(f"Resolved frame identifier: '{frame_col}', rebar column: '{rebar_col}', combo column: '{combo_col}'")
            # Parse records
            for _, row in df_results.iterrows():
                frame_name = str(row[frame_col]).strip()
                rebar_str = str(row[rebar_col]).strip()
                combo_name = str(row[combo_col]).strip() if combo_col else ""

                rebar_val = 0.0
                try:
                    val_str = rebar_str.replace('%', '').strip()
                    rebar_val = float(val_str)
                    # Convert to percent if retrieved as absolute ratio (e.g. 0.015 -> 1.5%)
                    if 'ratio' in rebar_col.lower() and 'rebar %' not in rebar_col.lower() and rebar_val < 0.1 and rebar_val > 0.0:
                        rebar_val *= 100.0
                except ValueError:
                    rebar_val = 0.0

                existing = rebar_data.get(frame_name)
                if existing is None or rebar_val > existing['pct']:
                    rebar_data[frame_name] = {'pct': rebar_val, 'combo': combo_name}
            print(f"Extracted design results for {len(rebar_data)} frames from PMM Envelope.")
        else:
            print("Failed to resolve frame/rebar columns in PMM Envelope table.")
    else:
        print("PMM Envelope table not resolved or empty.")
        
    return rebar_data


def read_etabs_table(sap_model, table_key):
    """
    Helper to query a specific ETABS database table into a pandas DataFrame.
    Dynamically scans returned COM tuple to find the table data CSV string.
    """
    try:
        ret = sap_model.DatabaseTables.GetTableForDisplayCSVString(TableKey=table_key, GroupName="")
        
        # If it returns a single string containing csv structure
        if isinstance(ret, str):
            if "Story" in ret or "Label" in ret or "," in ret:
                return pd.read_csv(io.StringIO(ret))
                
        # If it returns a tuple, scan for the CSV string dynamically
        if isinstance(ret, (tuple, list)):
            for val in ret:
                if isinstance(val, str) and ("," in val or "\n" in val or "Frame" in val or "Story" in val or "Grid" in val):
                    return pd.read_csv(io.StringIO(val))
    except Exception as e:
        print(f"Table '{table_key}' not readable: {e}")
    return None


def generate_pdf_report(x_grids, y_grids, aggregated_grids, file_path):
    """Generate a highly polished, premium multi-page PDF report on A3 Landscape paper."""
    print(f"Generating premium PDF report (A3 Landscape): '{file_path}'...")
    
    # Initialize A3 landscape document layout (margins = 54 pt = 0.75 in)
    doc = SimpleDocTemplate(
        file_path, 
        pagesize=landscape(A3),
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54
    )
    
    styles = getSampleStyleSheet()
    
    # Custom, visually stunning styling system
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
        alignment=1 # Centered
    )
    
    table_body_style = ParagraphStyle(
        'TableBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        textColor=colors.HexColor("#334155"),
        alignment=1 # Centered
    )
    
    table_body_bold_style = ParagraphStyle(
        'TableBodyBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        textColor=colors.HexColor("#1E293B"),
        alignment=1 # Centered
    )
    
    story = []
    
    # ------------------ PAGE 1: EXECUTIVE COVER PAGE ------------------
    story.append(Spacer(1, 100))
    story.append(Paragraph("COLUMN REBAR ANALYSIS REPORT", title_style))
    story.append(Paragraph("Programmatic Extraction & Visualization of Longitudinal Reinforcement from ETABS Models", subtitle_style))
    
    # Beautiful visual accent bar (Primary Theme Color: Deep Indigo #1E3A8A) (A3 drawable width = 1082 pt)
    accent_data = [['']]
    accent_table = Table(accent_data, colWidths=[1082], rowHeights=[6])
    accent_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#1E3A8A")),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(accent_table)
    story.append(Spacer(1, 60))
    
    # Project Metadata Block
    meta_title_style = ParagraphStyle('MetaTitle', fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor("#475569"))
    meta_val_style = ParagraphStyle('MetaVal', fontName='Helvetica', fontSize=12, textColor=colors.HexColor("#1E293B"))
    
    metadata = [
        [Paragraph("Project Name:", meta_title_style), Paragraph("Reinforcement Column Audit", meta_val_style)],
        [Paragraph("Analysis Software:", meta_title_style), Paragraph("CSI ETABS (Active API Interface)", meta_val_style)],
        [Paragraph("Date Generated:", meta_title_style), Paragraph(datetime.now().strftime("%B %d, %Y"), meta_val_style)],
        [Paragraph("Status:", meta_title_style), Paragraph("<font color='#10B981'><b>COMPLETED</b></font>", meta_val_style)]
    ]
    
    meta_table = Table(metadata, colWidths=[180, 902], rowHeights=[26]*4)
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
    
    # ------------------ PAGE 2: COLUMN KEY PLAN DRAWING ------------------
    story.append(Paragraph("Column Key Plan & Rebar Distribution", h1_style))
    story.append(Paragraph("The drawing below shows column geometries in plan (exaggerated scale for visibility) rotated based on local axes, alongside color-coded badges indicating maximum rebar percentage. Standard AABB collision-avoidance logic is implemented to eliminate overlap of label tags.", body_style))
    story.append(Spacer(1, 10))
    
    # Insert drawing flowable (Drawing dimensions: 1082 pt wide by 610 pt high on A3 Landscape)
    story.append(ColumnPlanFlowable(x_grids, y_grids, aggregated_grids, width=1082, height=600))
    story.append(PageBreak())
    
    # ------------------ PAGE 3: DETAILED REBAR DATA TABLE ------------------
    story.append(Paragraph("Detailed Reinforcement Data Table", h1_style))
    story.append(Paragraph("Summary of all column coordinates, section details, orientations, and the maximum rebar percentage with governing structural locations.", body_style))
    story.append(Spacer(1, 10))
    
    # Table headers
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
    
    # Add data rows
    for col in aggregated_grids:
        rebar = col['max_rebar']
        
        # Color coding cell rebar text based on percentage
        if rebar == 0.0:
            rebar_p = Paragraph("<font color='#64748B'><b>N/D</b></font>", table_body_bold_style)
        elif rebar < 2.7:
            rebar_p = Paragraph(f"<font color='#10B981'><b>{rebar:.2f}%</b></font>", table_body_bold_style)
        elif rebar <= 3.0:
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

    col_widths = [100, 140, 130, 70, 130, 70, 90, 120, 130, 102] # Sum = 1082 pt (A3 drawable width)
    
    # Construct ReportLab Table
    data_table = Table(table_rows, colWidths=col_widths, repeatRows=1)
    
    t_style = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1E3A8A")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
    ]
    
    # Zebra striping
    for idx in range(1, len(table_rows)):
        bg = colors.HexColor("#F8FAFC") if idx % 2 == 0 else colors.HexColor("#FFFFFF")
        t_style.append(('BACKGROUND', (0, idx), (-1, idx), bg))
        
    data_table.setStyle(TableStyle(t_style))
    story.append(data_table)
    
    # Build PDF
    doc.build(story, canvasmaker=NumberedCanvas)
    print("PDF report built successfully.")


def main():
    print("="*60)
    print("      ETABS COLUMN REBAR EXTRACTION & REPORT GENERATOR      ")
    print("="*60)
    
    sap_model = connect_to_etabs()
    if not sap_model:
        print("Fatal: Could not establish API link with ETABS.")
        sys.exit(1)
        
    try:
        # 1. Establish metric/SI units (kN, m, C) for predictable geometric coordinates
        original_units = sap_model.GetPresentUnits()
        sap_model.SetPresentUnits(6) # 6 = kN_m_C
        print("Present units set to: Kilonewtons, Meters, Celsius.")
        
        # 2. Read tables
        df_connectivity = read_etabs_table(sap_model, "Column Object Connectivity")
        df_column_bays = read_etabs_table(sap_model, "Column Bays")
        df_point_bays = read_etabs_table(sap_model, "Point Bays")
        df_grids = read_etabs_table(sap_model, "Grid Definitions - Grid Lines")
        
        # Check required geometry tables
        if any(df is None for df in [df_connectivity, df_column_bays, df_point_bays, df_grids]):
            print("Fatal error: Failed to read column mapping tables from ETABS database.")
            sap_model.SetPresentUnits(original_units)
            sys.exit(1)
            
        df_connectivity.columns = [c.strip() for c in df_connectivity.columns]
        df_column_bays.columns = [c.strip() for c in df_column_bays.columns]
        df_point_bays.columns = [c.strip() for c in df_point_bays.columns]
        df_grids.columns = [c.strip() for c in df_grids.columns]
        
        # 3. Parse Grid lines
        x_grids, y_grids = get_grid_lines(sap_model)
        
        # 4. Extract design results for rebar % directly from PMM Envelope table
        rebar_data = extract_rebar_design_results(sap_model)
        
        # 5. Build geometry mappings
        col_bay_to_pt_bay = {str(row['Label']).strip(): str(row['PointBayI']).strip() for _, row in df_column_bays.iterrows()}
        pt_bay_to_coord = {str(row['Label']).strip(): (float(row['X']), float(row['Y'])) for _, row in df_point_bays.iterrows()}
        
        # 6. Process columns from connectivity database table
        resolved_cols = []
        print("Processing column frames and mapping geometries...")
        for _, row in df_connectivity.iterrows():
            uniq_name = str(row['Unique Name']).strip()
            story = str(row['Story']).strip()
            col_bay = str(row['ColumnBay']).strip()
            
            # Look up point bay & coordinates
            pt_bay = col_bay_to_pt_bay.get(col_bay, None)
            if not pt_bay:
                continue
            coord = pt_bay_to_coord.get(pt_bay, None)
            if not coord:
                continue
                
            cx, cy = coord
            
            # Find closest grid intersection from the structural grids
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
            
            # Get section size name via API
            try:
                ret_sect = sap_model.FrameObj.GetSection(uniq_name)
                section_name = ret_sect[0]
            except Exception:
                section_name = "Default"
                
            # Get section sizes
            shape, w, d = get_section_dimensions(sap_model, section_name)
            
            # Get orientation angle
            angle = 0.0
            try:
                ret_axes = sap_model.FrameObj.GetLocalAxes(uniq_name)
                if ret_axes[-1] == 0:
                    angle = ret_axes[0]
            except Exception:
                pass
                
            # Get rebar percentage and governing PMM combo from PMM Envelope (with fallback to 0)
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
            print("No column frame elements resolved. Analysis aborted.")
            sap_model.SetPresentUnits(original_units)
            sys.exit(1)
            
        print(f"Resolved {len(resolved_cols)} columns (vertical frames) from connectivity.")
        
        # 7. Apply Layer 3 fallback (individual API queries) for frames with 0% rebar
        zero_rebar_cols = [c for c in resolved_cols if c['rebar_pct'] == 0.0]
        if len(zero_rebar_cols) > 0 and len(zero_rebar_cols) < len(resolved_cols):
            print(f"Layer 3 Fallback check: querying {len(zero_rebar_cols)} columns individually...")
            
            # Safe parsing function for numeric design outputs
            def get_floats(val):
                if isinstance(val, (list, tuple)):
                    res = []
                    for x in val:
                        try:
                            res.append(float(x))
                        except (ValueError, TypeError):
                            pass
                    return res
                else:
                    try:
                        return [float(val)]
                    except (ValueError, TypeError):
                        return []
                        
            for col in zero_rebar_cols:
                try:
                    # Query individual concrete design results
                    ret = sap_model.DesignConcrete.GetSummaryResultsColumn(col['frame_name'])
                    
                    if isinstance(ret, (tuple, list)) and len(ret) > 13:
                        num_res = ret[0]
                        top_area = ret[3]
                        bot_area = ret[4]
                        ret_code = ret[-1]
                        
                        if ret_code == 0 and num_res > 0:
                            top_floats = get_floats(top_area)
                            bot_floats = get_floats(bot_area)
                            
                            combined_floats = top_floats + bot_floats
                            if combined_floats:
                                max_area = max(combined_floats)
                                w = col['width']
                                d = col['depth']
                                shape = col['shape']
                                if shape == "Circular":
                                    sect_area = np.pi * (w / 2.0)**2
                                else:
                                    sect_area = w * d
                                    
                                pct = (max_area / sect_area) * 100.0
                                # Calibration check
                                if pct > 10.0 and sect_area < 10.0 and max_area > 1.0:
                                    sect_area_mm2 = sect_area * 1e6
                                    pct = (max_area / sect_area_mm2) * 100.0
                                col['rebar_pct'] = pct
                except Exception:
                    pass
                    
        # 8. Group columns by grid intersection coordinates
        column_coords = [(c['x'], c['y']) for c in resolved_cols]
        if not x_grids or not y_grids:
            print("Utilizing column coordinate clustering for virtual grids...")
            x_grids, y_grids = generate_virtual_grids(column_coords)
            
        grid_groups = {}
        for col in resolved_cols:
            key = (col['grid_label'], col['x'], col['y'])
            if key not in grid_groups:
                grid_groups[key] = []
            grid_groups[key].append(col)
            
        # 9. Aggregate grid point information
        aggregated_grids = []
        for (grid_label, gx, gy), cols in grid_groups.items():
            # Find column with max rebar
            max_rebar_col = max(cols, key=lambda c: c['rebar_pct'])
            max_rebar = max_rebar_col['rebar_pct']
            max_rebar_story = max_rebar_col['story_name']
            max_rebar_frame = max_rebar_col['frame_name']
            max_rebar_combo = max_rebar_col.get('pmm_combo', '')
            
            # Find the largest section size at this grid point to draw
            def get_area(col):
                w = col['width']
                d = col['depth']
                if col['shape'] == "Circular":
                    return np.pi * (w / 2.0)**2
                return w * d
                
            largest_col = max(cols, key=get_area)
            
            aggregated_grids.append({
                'grid_label': grid_label,
                'x': gx,
                'y': gy,
                'max_rebar': max_rebar,
                'max_rebar_story': max_rebar_story,
                'max_rebar_frame': max_rebar_frame,
                'max_rebar_combo': max_rebar_combo,
                'shape': largest_col['shape'],
                'width': largest_col['width'],
                'depth': largest_col['depth'],
                'angle': largest_col['angle'],
                'prop_name': largest_col['prop_name'],
                'all_frames': cols
            })
            
        # Sort aggregated grid list for the final data table
        aggregated_grids = sorted(aggregated_grids, key=lambda x: x['grid_label'])
        
        # 10. Build the PDF report with file-lock handling
        pdf_file = "Column_Rebar_Report.pdf"
        attempt = 1
        while True:
            try:
                generate_pdf_report(x_grids, y_grids, aggregated_grids, pdf_file)
                break
            except PermissionError:
                pdf_file = f"Column_Rebar_Report_{attempt}.pdf"
                print(f"Warning: PDF file is locked (possibly open in another viewer). Saving as '{pdf_file}'...")
                attempt += 1
        
        # Restore original ETABS units
        sap_model.SetPresentUnits(original_units)
        print("ETABS units restored to original status.")
        print("="*60)
        print("Success! The reinforcement layout analysis is finished.")
        print(f"Report location: {os.path.abspath(pdf_file)}")
        print("="*60)
        
    except Exception as e:
        print(f"An unexpected exception occurred during execution: {e}")
        try:
            sap_model.SetPresentUnits(original_units)
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()
