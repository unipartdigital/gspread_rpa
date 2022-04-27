
PYTHON ?= python3
PIPM := gspread openpyxl diff_match_patch gspread_formatting
PIPN := PIL,Pillow

.phony: clean gspread_rpa_demo

main: $(PIPM) $(PIPN) gspread_rpa_demo

upgrade:
	$(MAKE) upgrade=1 main

$(PIPM):
	@echo DEP: $@
ifeq ($(upgrade),)
	@$(PYTHON) -c "import $@" || $(PYTHON) -m pip install --user $@ ;
else
	$(PYTHON) -m pip install -U --user $@ ;
endif

$(PIPN):
	@echo DEP: $@
ifeq ($(upgrade),)
	@$(PYTHON) -c "import `echo $@ | cut -d, -f 1`" || \
$(PYTHON) -m pip install --user `echo $@ | cut -d, -f2` ;
else
	$(PYTHON) -m pip install -U --user `echo $@ | cut -d, -f2` ;
endif

gspread_rpa_demo:$(PIPM) $(PIPN) clean
	make -C src/gspread_rpa/demo

dist:
	$(PYTHON) -m build -n

localpip: dist
	$(PYTHON) -m pip download gspread-rpa -d localpip --only-binary=:all: -f ./dist/

clean:
	$(shell find . -type f -name "*~" -delete)
	$(shell find . -type f -name "*.log" -delete)
	$(shell find . -type f -name "*.log.[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]" -delete)
	$(shell find . -type f -name "*.pyc" -delete)
	$(shell find . -name "__pycache__" -delete)

distclean: clean
	$(shell find "gspread_rpa.egg-info" -type f -delete)
	$(shell find . -name "gspread_rpa.egg-info" -delete)
	$(shell find "dist" -type f -delete)
	$(shell find . -name "dist" -delete)

localpipclean:
	$(shell test -d localpip && find localpip -type f -delete)
	$(shell find . -name "localpip" -type d -delete)

help:
	@echo
	@echo "target"
	@echo "make clean"
	@echo "make main"
	@echo "dist"
	@echo "distclean"
	@echo "localpip"
	@echo "localpipclean"
	@echo "make upgrade"
	@echo "make $(PIPM)"
	@echo "make $(PIPN)"
