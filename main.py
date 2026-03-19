import sys
import os
import io
import fitz  # PyMuPDF
from PIL import Image
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QSpinBox,
                             QFileDialog, QMessageBox)
from PyQt6.QtGui import QPixmap, QImage, QPainter, QPen, QColor
from PyQt6.QtCore import Qt, pyqtSignal, QRect

# --- MODEL / CONTROLLER ---

class PrintProject:
    """Handles the core image buffer, PDF rendering, and aspect-correct slicing logic."""
    def __init__(self):
        self.source_image = None
        self.canvas_image = None  # Kept for compatibility; not eagerly built.
        self.rows = 1
        self.cols = 1
        self.overlap_mm = 10
        self.dpi = 300 # Standard print resolution
        self.canvas_w = 0
        self.canvas_h = 0
        self.scale_factor = 1.0
        self.scaled_w = 0
        self.scaled_h = 0
        self.offset_x = 0
        self.offset_y = 0
        
        # Standard A4 pixel dimensions at 300 DPI
        self.A4_W_PX = 2480
        self.A4_H_PX = 3508

    def load_file(self, filepath: str):
        if filepath.lower().endswith('.pdf'):
            doc = fitz.open(filepath)
            page = doc.load_page(0)
            pix = page.get_pixmap(dpi=self.dpi)
            mode = "RGBA" if pix.alpha else "RGB"
            self.source_image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        else:
            self.source_image = Image.open(filepath)
            
        self._rebuild_canvas()

    def update_grid(self, rows: int, cols: int, overlap_mm: int):
        self.rows = rows
        self.cols = cols
        self.overlap_mm = overlap_mm
        self._rebuild_canvas()

    def get_overlap_px(self) -> int:
        # Convert millimeters to pixels at 300 DPI
        return int(self.overlap_mm * (self.dpi / 25.4))

    def _rebuild_canvas(self):
        """Updates virtual canvas geometry and affine transform state (no pixel allocation)."""
        if not self.source_image:
            return

        overlap_px = self.get_overlap_px()
        # The distance between the start of one A4 page and the next
        step_x = self.A4_W_PX - overlap_px
        step_y = self.A4_H_PX - overlap_px

        # Total canvas size needed to hold all rows and columns with overlaps
        self.canvas_w = self.A4_W_PX + (self.cols - 1) * step_x
        self.canvas_h = self.A4_H_PX + (self.rows - 1) * step_y

        # 2. Calculate the scale factor to fit the image without distortion
        src_w, src_h = self.source_image.size
        self.scale_factor = min(self.canvas_w / src_w, self.canvas_h / src_h)
        self.scaled_w = int(src_w * self.scale_factor)
        self.scaled_h = int(src_h * self.scale_factor)

        # 4. Calculate coordinates to center the image on the canvas
        self.offset_x = (self.canvas_w - self.scaled_w) // 2
        self.offset_y = (self.canvas_h - self.scaled_h) // 2
        self.canvas_image = None

    def get_physical_size_mm(self) -> tuple[int, int]:
        # Physical size reflects the assembled poster, subtracting the hidden overlapped areas
        w = (self.cols * 210) - ((self.cols - 1) * self.overlap_mm)
        h = (self.rows * 297) - ((self.rows - 1) * self.overlap_mm)
        return w, h

    def get_slice(self, row: int, col: int) -> Image.Image:
        """Renders an A4 tile lazily from affine transform state."""
        if not self.source_image or self.canvas_w <= 0 or self.canvas_h <= 0:
            return None
        
        overlap_px = self.get_overlap_px()
        step_x = self.A4_W_PX - overlap_px
        step_y = self.A4_H_PX - overlap_px
        
        left = col * step_x
        upper = row * step_y
        right = left + self.A4_W_PX
        lower = upper + self.A4_H_PX

        tile = Image.new("RGBA", (self.A4_W_PX, self.A4_H_PX), (255, 255, 255, 255))

        img_left = self.offset_x
        img_top = self.offset_y
        img_right = self.offset_x + self.scaled_w
        img_bottom = self.offset_y + self.scaled_h

        inter_left = max(left, img_left)
        inter_top = max(upper, img_top)
        inter_right = min(right, img_right)
        inter_bottom = min(lower, img_bottom)

        if inter_left >= inter_right or inter_top >= inter_bottom:
            return tile

        src_left = (inter_left - self.offset_x) / self.scale_factor
        src_top = (inter_top - self.offset_y) / self.scale_factor
        src_right = (inter_right - self.offset_x) / self.scale_factor
        src_bottom = (inter_bottom - self.offset_y) / self.scale_factor

        sampled = self.source_image.crop((src_left, src_top, src_right, src_bottom))
        dest_w = inter_right - inter_left
        dest_h = inter_bottom - inter_top
        sampled = sampled.resize((dest_w, dest_h), Image.Resampling.LANCZOS)

        paste_x = inter_left - left
        paste_y = inter_top - upper
        mask = sampled if sampled.mode in ("RGBA", "LA") else None
        tile.paste(sampled, (paste_x, paste_y), mask)
        return tile

    def get_canvas_preview(self, max_width: int, max_height: int) -> Image.Image:
        """Returns a lightweight preview of the virtual canvas."""
        if not self.source_image or self.canvas_w <= 0 or self.canvas_h <= 0:
            return None

        ratio = min(max_width / self.canvas_w, max_height / self.canvas_h)
        ratio = min(1.0, ratio)
        preview_w = max(1, int(self.canvas_w * ratio))
        preview_h = max(1, int(self.canvas_h * ratio))

        preview = Image.new("RGBA", (preview_w, preview_h), (255, 255, 255, 255))
        draw_w = max(1, int(self.scaled_w * ratio))
        draw_h = max(1, int(self.scaled_h * ratio))
        draw_x = int(self.offset_x * ratio)
        draw_y = int(self.offset_y * ratio)

        resized = self.source_image.resize((draw_w, draw_h), Image.Resampling.LANCZOS)
        mask = resized if resized.mode in ("RGBA", "LA") else None
        preview.paste(resized, (draw_x, draw_y), mask)
        return preview

    def get_canvas_size(self) -> tuple[int, int]:
        return self.canvas_w, self.canvas_h

    def export_all(self, output_dir: str):
        if not self.source_image:
            return
            
        for r in range(self.rows):
            for c in range(self.cols):
                slice_img = self.get_slice(r, c)
                filename = f"slice_r{r+1}_c{c+1}.png"
                slice_img.save(os.path.join(output_dir, filename))

# --- VIEW & CONTROLLER HELPERS ---

def pil_to_qpixmap(pil_image: Image.Image) -> QPixmap:
    """Helper to safely convert PIL Image to PyQt6 QPixmap in memory."""
    bytes_io = io.BytesIO()
    pil_image.save(bytes_io, format="PNG")
    pixmap = QPixmap()
    pixmap.loadFromData(bytes_io.getvalue())
    return pixmap

class InteractiveGridPreview(QLabel):
    """A custom label that displays the full image with a grid and detects clicks."""
    slice_clicked = pyqtSignal(int, int)

    def __init__(self):
        super().__init__()
        self.rows = 1
        self.cols = 1
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setStyleSheet("border: 1px solid #aaa; background-color: #eee;")
        self.setMinimumSize(400, 400)

    def update_grid_params(self, rows: int, cols: int):
        self.rows = rows
        self.cols = cols

    def mousePressEvent(self, event):
        if not self.pixmap() or self.rows == 0 or self.cols == 0:
            return

        label_w, label_h = self.width(), self.height()
        pix_w, pix_h = self.pixmap().width(), self.pixmap().height()
        
        offset_x = (label_w - pix_w) // 2
        offset_y = (label_h - pix_h) // 2
        
        click_x = event.pos().x() - offset_x
        click_y = event.pos().y() - offset_y
        
        if 0 <= click_x <= pix_w and 0 <= click_y <= pix_h:
            col = int((click_x / pix_w) * self.cols)
            row = int((click_y / pix_h) * self.rows)
            
            col = max(0, min(col, self.cols - 1))
            row = max(0, min(row, self.rows - 1))
            
            self.slice_clicked.emit(row, col)

class SlicerMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.project = PrintProject()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("A4 Print Slicer")
        self.resize(1100, 700)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # Controls Layout
        controls = QHBoxLayout()
        
        self.btn_load = QPushButton("Load Image/PDF")
        self.btn_load.clicked.connect(self.load_file)
        controls.addWidget(self.btn_load)

        controls.addWidget(QLabel("Rows:"))
        self.spin_rows = QSpinBox()
        self.spin_rows.setRange(1, 20)
        self.spin_rows.valueChanged.connect(self.update_grid)
        controls.addWidget(self.spin_rows)

        controls.addWidget(QLabel("Cols:"))
        self.spin_cols = QSpinBox()
        self.spin_cols.setRange(1, 20)
        self.spin_cols.valueChanged.connect(self.update_grid)
        controls.addWidget(self.spin_cols)

        controls.addWidget(QLabel("Overlap:"))
        self.spin_overlap = QSpinBox()
        self.spin_overlap.setRange(0, 50)
        self.spin_overlap.setSuffix(" mm")
        self.spin_overlap.setValue(10) # Default 10mm overlap
        self.spin_overlap.valueChanged.connect(self.update_grid)
        controls.addWidget(self.spin_overlap)

        self.lbl_size = QLabel("Assembled Size: 210mm x 297mm")
        controls.addWidget(self.lbl_size)

        self.btn_export = QPushButton("Export Slices")
        self.btn_export.clicked.connect(self.export_slices)
        controls.addWidget(self.btn_export)

        layout.addLayout(controls)

        # Preview Layout
        previews = QHBoxLayout()
        
        # Left: Full image with interactive grid
        self.full_preview = InteractiveGridPreview()
        self.full_preview.slice_clicked.connect(self.preview_slice)
        previews.addWidget(self.full_preview, stretch=1)

        # Right: Individual slice preview
        self.lbl_slice_preview = QLabel("Select a slice to preview")
        self.lbl_slice_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_slice_preview.setStyleSheet("border: 1px solid #ccc; background-color: white;")
        previews.addWidget(self.lbl_slice_preview, stretch=1)

        layout.addLayout(previews)

    def load_file(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Select Image or PDF", "", "Media Files (*.png *.jpg *.jpeg *.pdf)"
        )
        if filepath:
            self.project.load_file(filepath)
            self.update_grid()

    def update_grid(self):
        if not self.project.source_image:
            return

        rows = self.spin_rows.value()
        cols = self.spin_cols.value()
        overlap = self.spin_overlap.value()

        self.project.update_grid(rows, cols, overlap)
        
        w_mm, h_mm = self.project.get_physical_size_mm()
        self.lbl_size.setText(f"Assembled Size: {w_mm}mm x {h_mm}mm")
        self.full_preview.update_grid_params(rows, cols)

        preview_img = self.project.get_canvas_preview(500, 500)
        canvas_pixmap = pil_to_qpixmap(preview_img)
        scaled_canvas = canvas_pixmap

        painter = QPainter(scaled_canvas)
        pen = QPen(QColor(255, 0, 0, 150))
        pen.setWidth(2)
        painter.setPen(pen)

        # Calculate the exact pixel boundaries for the red boxes on the scaled preview
        canvas_w, canvas_h = self.project.get_canvas_size()
        scale_x = scaled_canvas.width() / canvas_w
        scale_y = scaled_canvas.height() / canvas_h
        
        overlap_px = self.project.get_overlap_px()
        step_x = self.project.A4_W_PX - overlap_px
        step_y = self.project.A4_H_PX - overlap_px

        # Draw a rectangle for every A4 sheet
        for r in range(rows):
            for c in range(cols):
                left = int(c * step_x * scale_x)
                upper = int(r * step_y * scale_y)
                width = int(self.project.A4_W_PX * scale_x)
                height = int(self.project.A4_H_PX * scale_y)
                painter.drawRect(QRect(left, upper, width, height))

        painter.end()

        self.full_preview.setPixmap(scaled_canvas)
        self.preview_slice(0, 0)

    def preview_slice(self, row: int, col: int):
        slice_img = self.project.get_slice(row, col)
        if slice_img:
            pixmap = pil_to_qpixmap(slice_img)
            scaled_pixmap = pixmap.scaled(
                500, 500, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
            )
            self.lbl_slice_preview.setPixmap(scaled_pixmap)

    def export_slices(self):
        if not self.project.source_image:
            QMessageBox.warning(self, "Error", "Load a file first.")
            return

        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.project.export_all(folder)
            QMessageBox.information(self, "Success", f"Exported {self.project.rows * self.project.cols} files.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SlicerMainWindow()
    window.show()
    sys.exit(app.exec())