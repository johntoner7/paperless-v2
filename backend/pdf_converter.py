import logging
import os
import subprocess
import time

LOG = logging.getLogger(__name__)

_BIN = os.environ.get('PDF2HTML_BIN', 'pdf2htmlEX')


def convert(input_pdf: str, output_dir: str, timeout: int = 120) -> dict:
    """Run pdf2htmlEX on *input_pdf*, writing a single self-contained HTML file to *output_dir*.

    Returns:
        {
            "html_path":    str | None,   # absolute path to the output HTML, or None on failure
            "duration_sec": float,
            "success":      bool,
            "exit_code":    int | None,
            "stderr":       str,
        }
    """
    os.makedirs(output_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(input_pdf))[0]
    html_filename = stem + '.html'

    cmd = [
        _BIN,
        '--dest-dir', output_dir,
        # Embed everything so the output is a single portable HTML file.
        '--embed-css', '1',
        '--embed-font', '1',
        '--embed-image', '1',
        '--embed-javascript', '1',
        '--embed-outline', '1',
        '--process-outline', '0',
        input_pdf,
        html_filename,
    ]

    LOG.info('Running: %s', ' '.join(cmd))
    start = time.monotonic()

    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        duration = round(time.monotonic() - start, 2)
        LOG.error('pdf2htmlEX timed out after %ds', timeout)
        return {
            'html_path': None,
            'duration_sec': duration,
            'success': False,
            'exit_code': None,
            'stderr': f'Timed out after {timeout}s',
        }
    except FileNotFoundError:
        return {
            'html_path': None,
            'duration_sec': 0.0,
            'success': False,
            'exit_code': None,
            'stderr': f'pdf2htmlEX binary not found (PDF2HTML_BIN={_BIN!r})',
        }

    duration = round(time.monotonic() - start, 2)
    stderr_text = proc.stderr.decode('utf-8', errors='replace')
    success = proc.returncode == 0

    if not success:
        LOG.error('pdf2htmlEX exit %d: %s', proc.returncode, stderr_text)

    return {
        'html_path': os.path.join(output_dir, html_filename) if success else None,
        'duration_sec': duration,
        'success': success,
        'exit_code': proc.returncode,
        'stderr': stderr_text,
    }
