PYTHON ?= python

.PHONY: all check docx figures

all: docx

check:
	$(PYTHON) tools/build.py check

docx:
	$(PYTHON) tools/build.py docx

figures:
	$(PYTHON) tools/draw_search_pipeline.py
