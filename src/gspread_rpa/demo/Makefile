PYTHON ?= python3
TMP := $(shell $(PYTHON) -c 'import tempfile ; print (tempfile.gettempdir())')

%.py.done : %.py
	$(PYTHON) $< && touch $@

all: $(addsuffix .done, $(wildcard *.py))

clean:
	$(RM) *.log
	$(RM) *.log.[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]
	$(RM) *.done
	$(RM) $(TMP)/[0-9][0-9]-demo-revision-[0-9][0-9]*.*
