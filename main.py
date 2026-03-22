import sys
import os
import io
import fitz  # PyMuPDF
from PIL import Image
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QSpinBox,
                             QFileDialog, QMessageBox, QProgressDialog,
                             QDialog, QCheckBox, QDoubleSpinBox, QDialogButtonBox)
from PyQt6.QtGui import QPixmap, QImage, QPainter, QPen, QColor
from PyQt6.QtCore import Qt, pyqtSignal, QRect

# --- MODEL / CONTROLLER ---

class PrintProject:
    """Handles the core image buffer, PDF rendering, and aspect-correct slicing logic."""
    def __init__(self):
        self.source_image = None
        self.proxy_image = None
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

        self.proxy_image = self._build_proxy_image(self.source_image)
            
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

        # Keep source-derived geometry values for compatibility.
        self.scale_factor, self.scaled_w, self.scaled_h, self.offset_x, self.offset_y = self._compute_transform(
            self.source_image
        )
        self.canvas_image = None

    def _compute_transform(self, image: Image.Image) -> tuple[float, int, int, int, int]:
        """Compute fit-and-center transform for any image against current canvas."""
        src_w, src_h = image.size
        scale_factor = min(self.canvas_w / src_w, self.canvas_h / src_h)
        scaled_w = int(src_w * scale_factor)
        scaled_h = int(src_h * scale_factor)
        offset_x = (self.canvas_w - scaled_w) // 2
        offset_y = (self.canvas_h - scaled_h) // 2
        return scale_factor, scaled_w, scaled_h, offset_x, offset_y

    def _build_proxy_image(self, image: Image.Image) -> Image.Image:
        """Build an in-memory lightweight proxy, capped to 1080p bounds."""
        max_w = 1920
        max_h = 1080
        src_w, src_h = image.size
        ratio = min(max_w / src_w, max_h / src_h, 1.0)
        proxy_w = max(1, int(src_w * ratio))
        proxy_h = max(1, int(src_h * ratio))
        if proxy_w == src_w and proxy_h == src_h:
            return image.copy()
        return image.resize((proxy_w, proxy_h), Image.Resampling.LANCZOS)

    def get_physical_size_mm(self) -> tuple[int, int]:
        # Physical size reflects the assembled poster, subtracting the hidden overlapped areas
        w = (self.cols * 210) - ((self.cols - 1) * self.overlap_mm)
        h = (self.rows * 297) - ((self.rows - 1) * self.overlap_mm)
        return w, h

    def get_slice(self, row: int, col: int, use_proxy: bool = False) -> Image.Image:
        """Renders an A4 tile lazily from affine transform state."""
        image = self.proxy_image if use_proxy else self.source_image
        if not image or self.canvas_w <= 0 or self.canvas_h <= 0:
            return None
        
        overlap_px = self.get_overlap_px()
        step_x = self.A4_W_PX - overlap_px
        step_y = self.A4_H_PX - overlap_px
        
        left = col * step_x
        upper = row * step_y
        right = left + self.A4_W_PX
        lower = upper + self.A4_H_PX

        tile = Image.new("RGBA", (self.A4_W_PX, self.A4_H_PX), (255, 255, 255, 255))

        scale_factor, scaled_w, scaled_h, offset_x, offset_y = self._compute_transform(image)
        img_left = offset_x
        img_top = offset_y
        img_right = offset_x + scaled_w
        img_bottom = offset_y + scaled_h

        inter_left = max(left, img_left)
        inter_top = max(upper, img_top)
        inter_right = min(right, img_right)
        inter_bottom = min(lower, img_bottom)

        if inter_left >= inter_right or inter_top >= inter_bottom:
            return tile

        src_left = (inter_left - offset_x) / scale_factor
        src_top = (inter_top - offset_y) / scale_factor
        src_right = (inter_right - offset_x) / scale_factor
        src_bottom = (inter_bottom - offset_y) / scale_factor

        sampled = image.crop((src_left, src_top, src_right, src_bottom))
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
        if not self.proxy_image or self.canvas_w <= 0 or self.canvas_h <= 0:
            return None

        ratio = min(max_width / self.canvas_w, max_height / self.canvas_h)
        ratio = min(1.0, ratio)
        preview_w = max(1, int(self.canvas_w * ratio))
        preview_h = max(1, int(self.canvas_h * ratio))

        preview = Image.new("RGBA", (preview_w, preview_h), (255, 255, 255, 255))
        _, scaled_w, scaled_h, offset_x, offset_y = self._compute_transform(self.proxy_image)
        draw_w = max(1, int(scaled_w * ratio))
        draw_h = max(1, int(scaled_h * ratio))
        draw_x = int(offset_x * ratio)
        draw_y = int(offset_y * ratio)

        resized = self.proxy_image.resize((draw_w, draw_h), Image.Resampling.LANCZOS)
        mask = resized if resized.mode in ("RGBA", "LA") else None
        preview.paste(resized, (draw_x, draw_y), mask)
        return preview

    def get_canvas_size(self) -> tuple[int, int]:
        return self.canvas_w, self.canvas_h

    def export_all(self, output_dir: str, progress_callback=None, cancel_check=None) -> tuple[int, bool]:
        if not self.source_image:
            return 0, False
        
        exported = 0
        total = self.rows * self.cols
            
        for r in range(self.rows):
            for c in range(self.cols):
                if cancel_check and cancel_check():
                    return exported, True
                slice_img = self.get_slice(r, c)
                filename = f"slice_r{r+1}_c{c+1}.png"
                slice_img.save(os.path.join(output_dir, filename))
                exported += 1
                if progress_callback:
                    progress_callback(exported, total)
        return exported, False

    def export_pdf(
        self,
        output_pdf: str,
        image_format: str = "jpeg",
        jpeg_quality: int = 80,
        downscale: float = 0.75,
        progress_callback=None,
        cancel_check=None,
    ) -> tuple[int, bool]:
        if not self.source_image:
            return 0, False

        image_format = (image_format or "").lower().strip()
        if image_format not in ("jpeg", "jpg", "png"):
            image_format = "jpeg"

        exported = 0
        total = self.rows * self.cols
        pdf_doc = fitz.open()
        a4_rect = fitz.paper_rect("a4")

        for r in range(self.rows):
            for c in range(self.cols):
                if cancel_check and cancel_check():
                    pdf_doc.close()
                    return exported, True

                slice_img = self.get_slice(r, c)

                # Reduce pixel dimensions before embedding to keep the PDF size down.
                if downscale and downscale < 0.999:
                    new_w = max(1, int(slice_img.width * downscale))
                    new_h = max(1, int(slice_img.height * downscale))
                    slice_img = slice_img.resize((new_w, new_h), Image.Resampling.LANCZOS)

                img_buffer = io.BytesIO()
                if image_format in ("jpeg", "jpg"):
                    slice_img = slice_img.convert("RGB")
                    slice_img.save(
                        img_buffer,
                        format="JPEG",
                        quality=max(1, min(int(jpeg_quality), 95)),
                        optimize=True,
                    )
                else:
                    slice_img = slice_img.convert("RGB")
                    slice_img.save(img_buffer, format="PNG", optimize=True)

                page = pdf_doc.new_page(width=a4_rect.width, height=a4_rect.height)
                page.insert_image(page.rect, stream=img_buffer.getvalue(), keep_proportion=False)

                exported += 1
                if progress_callback:
                    progress_callback(exported, total)

        pdf_doc.save(output_pdf)
        pdf_doc.close()
        return exported, False

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

class PdfExportOptionsDialog(QDialog):
    def __init__(self, project: PrintProject, parent=None):
        super().__init__(parent)
        self.project = project
        self.total = project.rows * project.cols

        self.setWindowTitle("PDF Export Options")
        self.setMinimumWidth(420)

        root = QVBoxLayout(self)

        self.chk_jpeg = QCheckBox("Use JPEG compression (recommended, smaller)")
        self.chk_jpeg.setChecked(True)
        root.addWidget(self.chk_jpeg)

        row_quality = QHBoxLayout()
        row_quality.addWidget(QLabel("JPEG quality:"))
        self.spin_quality = QSpinBox()
        self.spin_quality.setRange(1, 95)
        self.spin_quality.setValue(80)
        row_quality.addWidget(self.spin_quality)
        root.addLayout(row_quality)

        row_scale = QHBoxLayout()
        row_scale.addWidget(QLabel("Export scale:"))
        self.spin_scale = QDoubleSpinBox()
        self.spin_scale.setRange(0.2, 1.0)
        self.spin_scale.setSingleStep(0.05)
        self.spin_scale.setDecimals(2)
        self.spin_scale.setValue(0.75)
        row_scale.addWidget(self.spin_scale)
        root.addLayout(row_scale)

        self.lbl_estimate = QLabel("Estimated PDF size: (not calculated yet)")
        self.lbl_estimate.setWordWrap(True)
        root.addWidget(self.lbl_estimate)

        self.btn_estimate = QPushButton("Estimate size")
        root.addWidget(self.btn_estimate)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self.chk_jpeg.toggled.connect(self._sync_controls)
        self._sync_controls()
        self.btn_estimate.clicked.connect(self.compute_estimate)

        # Initial estimate so the user sees an immediate size preview.
        self.compute_estimate()

    def _sync_controls(self):
        self.spin_quality.setEnabled(self.chk_jpeg.isChecked())

    def get_options(self) -> tuple[str, int, float]:
        image_format = "jpeg" if self.chk_jpeg.isChecked() else "png"
        return image_format, int(self.spin_quality.value()), float(self.spin_scale.value())

    def compute_estimate(self):
        if not self.project.source_image:
            self.lbl_estimate.setText("Estimated PDF size: load a file first.")
            return

        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

            # Use proxy for speed; estimate is approximate.
            slice_img = self.project.get_slice(0, 0, use_proxy=True)

            scale = self.spin_scale.value()
            if scale and scale < 0.999:
                new_w = max(1, int(slice_img.width * scale))
                new_h = max(1, int(slice_img.height * scale))
                slice_img = slice_img.resize((new_w, new_h), Image.Resampling.LANCZOS)

            img_buffer = io.BytesIO()
            if self.chk_jpeg.isChecked():
                q = int(self.spin_quality.value())
                slice_img = slice_img.convert("RGB")
                slice_img.save(
                    img_buffer,
                    format="JPEG",
                    quality=max(1, min(q, 95)),
                    optimize=True,
                )
                fmt_label = f"JPEG (q={q})"
            else:
                slice_img = slice_img.convert("RGB")
                slice_img.save(img_buffer, format="PNG", optimize=True)
                fmt_label = "PNG"

            bytes_per_slice = len(img_buffer.getvalue())
            overhead_factor = 1.08  # Rough PDF object/page overhead.
            estimate_bytes = bytes_per_slice * self.total * overhead_factor
            estimate_mb = estimate_bytes / (1024 * 1024)

            self.lbl_estimate.setText(
                f"Estimated PDF size (~{fmt_label}, ~{self.total} pages): {estimate_mb:.1f} MB"
            )
        finally:
            QApplication.restoreOverrideCursor()

class SlicerMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.project = PrintProject()
        self.current_row = 0
        self.current_col = 0
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
        self.current_row = max(0, min(self.current_row, rows - 1))
        self.current_col = max(0, min(self.current_col, cols - 1))
        self._render_full_preview()
        self.preview_slice(self.current_row, self.current_col)

    def _render_full_preview(self):
        rows = self.spin_rows.value()
        cols = self.spin_cols.value()
        preview_img = self.project.get_canvas_preview(500, 500)
        canvas_pixmap = pil_to_qpixmap(preview_img)
        scaled_canvas = canvas_pixmap

        painter = QPainter(scaled_canvas)

        canvas_w, canvas_h = self.project.get_canvas_size()
        scale_x = scaled_canvas.width() / canvas_w
        scale_y = scaled_canvas.height() / canvas_h
        
        overlap_px = self.project.get_overlap_px()
        step_x = self.project.A4_W_PX - overlap_px
        step_y = self.project.A4_H_PX - overlap_px

        # Draw a rectangle for every A4 sheet and highlight the active one.
        for r in range(rows):
            for c in range(cols):
                left = int(c * step_x * scale_x)
                upper = int(r * step_y * scale_y)
                width = int(self.project.A4_W_PX * scale_x)
                height = int(self.project.A4_H_PX * scale_y)
                if r == self.current_row and c == self.current_col:
                    painter.fillRect(QRect(left, upper, width, height), QColor(255, 215, 0, 70))
                    active_pen = QPen(QColor(255, 215, 0, 230))
                    active_pen.setWidth(3)
                    painter.setPen(active_pen)
                    painter.drawRect(QRect(left, upper, width, height))
                grid_pen = QPen(QColor(255, 0, 0, 150))
                grid_pen.setWidth(2)
                painter.setPen(grid_pen)
                painter.drawRect(QRect(left, upper, width, height))

        painter.end()

        self.full_preview.setPixmap(scaled_canvas)

    def preview_slice(self, row: int, col: int):
        self.current_row = row
        self.current_col = col
        slice_img = self.project.get_slice(row, col, use_proxy=True)
        if slice_img:
            pixmap = pil_to_qpixmap(slice_img)
            scaled_pixmap = pixmap.scaled(
                500, 500, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
            )
            self.lbl_slice_preview.setPixmap(scaled_pixmap)
            self._render_full_preview()

    def export_slices(self):
        if not self.project.source_image:
            QMessageBox.warning(self, "Error", "Load a file first.")
            return

        format_dialog = QMessageBox(self)
        format_dialog.setWindowTitle("Export Format")
        format_dialog.setText("Choose export format:")
        png_btn = format_dialog.addButton("PNG Slices", QMessageBox.ButtonRole.AcceptRole)
        pdf_btn = format_dialog.addButton("Single PDF", QMessageBox.ButtonRole.AcceptRole)
        cancel_btn = format_dialog.addButton(QMessageBox.StandardButton.Cancel)
        format_dialog.exec()

        clicked = format_dialog.clickedButton()
        if clicked == cancel_btn:
            return

        total = self.project.rows * self.project.cols

        def run_progress_export(export_fn):
            progress = QProgressDialog("Exporting...", "Cancel", 0, total, self)
            progress.setWindowTitle("Export Progress")
            progress.setWindowModality(Qt.WindowModality.WindowModal)
            progress.setMinimumDuration(0)
            progress.setValue(0)

            def on_progress(done: int, total_items: int):
                progress.setLabelText(f"Exporting... ({done}/{total_items})")
                progress.setValue(done)
                QApplication.processEvents()

            exported, canceled = export_fn(progress, on_progress)
            progress.close()
            return exported, canceled

        exported = 0
        canceled = False

        if clicked == png_btn:
            folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
            if not folder:
                return

            exported, canceled = run_progress_export(
                lambda progress, on_progress: self.project.export_all(
                    folder,
                    progress_callback=on_progress,
                    cancel_check=progress.wasCanceled,
                )
            )

        elif clicked == pdf_btn:
            options_dialog = PdfExportOptionsDialog(self.project, self)
            if options_dialog.exec() != QDialog.DialogCode.Accepted:
                return

            image_format, jpeg_quality, downscale = options_dialog.get_options()

            pdf_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Slices as PDF",
                "slices.pdf",
                "PDF Files (*.pdf)",
            )
            if not pdf_path:
                return
            if not pdf_path.lower().endswith(".pdf"):
                pdf_path += ".pdf"

            def do_export(progress, on_progress):
                return self.project.export_pdf(
                    pdf_path,
                    image_format=image_format,
                    jpeg_quality=jpeg_quality,
                    downscale=downscale,
                    progress_callback=on_progress,
                    cancel_check=progress.wasCanceled,
                )

            exported, canceled = run_progress_export(do_export)

        if canceled:
            QMessageBox.information(self, "Export Canceled", f"Export canceled after {exported} pages.")
        elif clicked == png_btn:
            QMessageBox.information(self, "Success", f"Exported {exported} PNG files.")
        elif clicked == pdf_btn:
            QMessageBox.information(self, "Success", f"Exported {exported} pages to PDF.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SlicerMainWindow()
    window.show()
    sys.exit(app.exec())