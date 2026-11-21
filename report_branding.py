"""Self-contained FINDEX report branding and standard vector icons."""
import base64
import html
from pathlib import Path


def report_logo_uri(app_root):
    path = Path(app_root) / 'resources' / 'branding' / 'report_logo.png'
    return 'data:image/png;base64,' + base64.b64encode(path.read_bytes()).decode('ascii') if path.is_file() else ''


def report_icon(name):
    shapes = {
        'file': '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6M8 13h8M8 17h6"/>',
        'folder': '<path d="M3 7V5a2 2 0 0 1 2-2h5l2 3h7a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/>',
        'search': '<circle cx="10.5" cy="10.5" r="7.5"/><path d="m16 16 5 5"/>',
        'download': '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M12 3v12m-5-5 5 5 5-5"/>',
        'printer': '<path d="M6 9V3h12v6M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><path d="M6 14h12v8H6zM18 12h.01"/>',
        'sun': '<circle cx="12" cy="12" r="4"/><path d="M12 2v2m0 16v2M2 12h2m16 0h2M4.9 4.9l1.4 1.4m11.4 11.4 1.4 1.4M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
        'moon': '<path d="M20.9 13A9 9 0 0 1 11 3.1 9 9 0 1 0 20.9 13Z"/>',
    }
    return ('<svg class="report-icon" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" '
            'width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" '
            'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false">'
            + shapes[name] + '</svg>')


def report_brand_header(logo_uri, name, subtitle, version, context='', heading=False):
    logo = f'<img class="report-logo" src="{logo_uri}" alt="{html.escape(name)} logo">' if logo_uri else report_icon('file')
    context_html = f'<div class="report-context">{html.escape(context)}</div>' if context else ''
    title_tag = 'h1' if heading else 'div'
    return (f'<div class="report-brand"><span class="report-logo-tile">{logo}</span>'
            f'<div><{title_tag} class="report-brand-name">{html.escape(name)} '
            f'<span class="report-version">v{html.escape(version)}</span></{title_tag}>'
            f'<div class="report-brand-subtitle">{html.escape(subtitle)}</div>{context_html}</div></div>')


def report_head_tags(logo_uri, name):
    favicon = f'<link rel="icon" type="image/png" href="{logo_uri}">' if logo_uri else ''
    return f'<meta name="application-name" content="{html.escape(name)}">' + favicon


REPORT_BRAND_CSS = """
    .report-brand { display:flex; align-items:center; gap:10px; margin-bottom:10px; color:inherit; }
    .report-logo-tile { display:inline-flex; align-items:center; justify-content:center;
        width:40px; height:40px; flex:0 0 40px; background:#111625; border-radius:6px;
        -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .report-logo { width:36px; height:36px; object-fit:contain; }
    .report-brand-name { margin:0; font-size:20px; line-height:1.2; font-weight:700; letter-spacing:.3px; }
    .report-brand-subtitle, .report-context { font-size:12px; line-height:1.4; color:var(--text-muted,#475569); }
    .report-version { font-size:11px; font-weight:500; white-space:nowrap; }
    .report-icon { display:inline-block; vertical-align:-3px; flex-shrink:0; }
    .brand-title .report-brand { margin-bottom:0; }
    .report-footer { margin-top:20px; padding-top:10px; border-top:1px solid var(--border-subtle,#e2e8f0);
        color:var(--text-muted,#475569); font-size:11px; text-align:center; }
    body .report-footer { color:var(--text-muted,#475569); }
    .theme-light-icon, .theme-dark-icon { display:inline-flex; }
    [data-theme="light"] .theme-light-icon, [data-theme="dark"] .theme-dark-icon { display:none; }
    @media print {
        .report-brand { margin-bottom:6px; break-inside:avoid; }
        .report-logo-tile { width:30px; height:30px; flex-basis:30px; }
        .report-logo { width:28px; height:28px; }
        .report-brand-name { font-size:12pt; }
        .report-brand-subtitle, .report-context { font-size:8pt; color:#475569; }
    }
"""
