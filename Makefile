-include .env

PYTHON?=python
ifeq ($(strip $(VIRTUAL_ENV)),)
VENV=source .venv/Scripts/activate;
endif
PYTHON:=$(VENV) $(PYTHON)
KARTVERKTYG=$(PYTHON) src/platsplanering.py
SCHEMAVERKTYG=$(PYTHON) src/schema.py
UPLOADSCRIPT=$(PYTHON) src/uploadfiles.py
SENDMAILSCRIPT=$(PYTHON) src/sendemail.py
STAGE=stage
CLUBNAME?=ESS
UV=$(if $(CI),,uv)

dirs=$(STAGE)

$(dirs):
	mkdir -p $@

.venv:
	$(PYTHON) -m venv $@

$(STAGE)/requirements.txt: requirements.txt
	$(if $(CI),,$(VENV)) $(UV) pip install -r $<
	touch $@

.phony: prereqs
prereqs: $(dirs) $(if $(CI),,.venv) $(STAGE)/requirements.txt

.phony: karta
karta: prereqs
	$(KARTVERKTYG) \
		$(if $(REQUEST_SOURCE),--requests $(REQUEST_SOURCE)) \
		$(if $(EXMEMBERS),--exmembers $(EXMEMBERS)) \
		$(if $(ONLAND),--onland $(ONLAND))

.phony: schema
schema: prereqs
	$(SCHEMAVERKTYG) \
		--outdir $(STAGE) \
		--header "Schema $(CLUBNAME)" \
		--driversheetid $(DRIVERSCHEDULE) \
		--mapfile "varvskarta*.pptx"

# emails: SHEET_ID=$(EMAIL_SHEET_ID)
# emails: prereqs
# 	$(PYTHON) src/generate_email.py  --sheetid $(SHEET_ID)

sendmail: prereqs
	$(SENDMAILSCRIPT) \
		--receiver $(EMAIL_RECEIVER) \
		--template templates/email-template.html \
		--replacement "varvschef=$(VARVSCHEF)"

upload: prereqs
	$(UPLOADSCRIPT)
