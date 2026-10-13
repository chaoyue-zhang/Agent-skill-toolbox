.PHONY: all build-html open clean

PY ?= python3

all: build-html

build-html:
	$(PY) scripts/build_static_html.py

open:
	$(PY) -c "import webbrowser,os;webbrowser.open('file://'+os.path.abspath('docs/index.html'))"

clean:
	rm -rf docs/index.html docs/assets
