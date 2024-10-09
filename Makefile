PYTHON?=python
ifeq ($(strip $(VIRTUAL_ENV)),)
VENV=source .venv/Scripts/activate;
endif
PYTHON:=$(VENV) $(PYTHON)
KARTVERKTYG=$(PYTHON) src/platsplanering.py
SCHEMAVERKTYG=$(PYTHON) src/main.py
STAGE=stage
CLUBNAME?=ESS

dirs=$(STAGE)

$(dirs):
	mkdir -p $@

.venv:
	$(PYTHON) -m venv $@

$(STAGE)/requirements.txt: requirements.txt
	$(VENV) pip install -r $<
	touch $@

.phony: prereqs
prereqs: $(dirs) .venv $(STAGE)/requirements.txt

.phony: karta
karta: prereqs
	$(KARTVERKTYG)

.phony: schema
schema: prereqs
	$(SCHEMAVERKTYG) \
		--outdir $(STAGE) \
		--template templates/template.xlsx \
		--header "Schema $(CLUBNAME)" \
		--mapfile "varvskarta*.pptx"
