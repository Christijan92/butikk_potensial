"""Butikkpotensial-analyse for sammenligning av varegrupper."""

if __package__ in (None, ""):
    # Support direct execution of __init__.py for development and quick checks.
    from analysis import AnalysisResult, analyze_files, export_analysis
else:
    from .analysis import AnalysisResult, analyze_files, export_analysis

__all__ = ["AnalysisResult", "analyze_files", "export_analysis"]
