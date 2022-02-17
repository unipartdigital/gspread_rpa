PYTHON ?= python3

%.py.done : %.py
	$(PYTHON) $< && touch $@

all: $(addsuffix .done, $(wildcard *.py))

clean:
	$(RM) *.log
	$(RM) *.log.[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]
	$(RM) *.done
