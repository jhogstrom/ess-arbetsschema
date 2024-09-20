PYTHON?=python
KARTVERKTYG=$(PYTHON) src/platsplanering.py
SCHEMAVERKTYG=$(PYTHON) src/main.py
STAGE=stage
CLUBNAME?=ESS

dirs=$(STAGE)

$(dirs):
	mkdir -p $@


karta: $(STAGE)
	$(KARTVERKTYG)

schema: $(STAGE)
	$(SCHEMAVERKTYG) \
		--outdir $(STAGE) \
		--template templates/template.xlsx \
		--header "Schema $(CLUBNAME)" \
		--mapfile "varvskarta*.pptx"
