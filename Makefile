PYTHON?=python
KARTVERKTYG=$(PYTHON) src/platsplanering.py
SCHEMAVERKTYG=$(PYTHON) src/main.py
STAGE=stage

dirs=$(STAGE)

$(dirs):
	mkdir -p $@


karta: $(STAGE)
	$(KARTVERKTYG)

schema: $(STAGE)
	$(SCHEMAVERKTYG)
