"""追記ログ列のルール処理（分割・日時・記入者・識別子・マスク・用語集・時系列の描画）。純粋関数のみ。"""
from logproc.glossary import GlossaryHit, apply_glossary, glossary_hits
from logproc.mask import MaskSpan, find_mask_spans, mask_text
from logproc.models import AuthorInfo, LogParse, Segment, SplitOptions, WhenInfo
from logproc.people import Person, PeopleIndex
from logproc.render import format_author, format_when, render_timeline, review_notes, timeline_order
from logproc.segment import parse_log

__all__ = [
    "AuthorInfo", "GlossaryHit", "LogParse", "MaskSpan", "PeopleIndex", "Person", "Segment", "SplitOptions", "WhenInfo",
    "apply_glossary", "find_mask_spans", "format_author", "format_when", "glossary_hits", "mask_text", "parse_log",
    "render_timeline", "review_notes", "timeline_order",
]
