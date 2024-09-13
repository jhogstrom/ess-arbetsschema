# Arbetsschema

Detta program genererar arbetsschema och sjösättningsschema från en rapport FRÅN BAS.

Rapporten läses in, ett set av datum för året skapas, varefter en rapport genereras per
dag definierad som sjösättningsdag, såväl arbetspass som sjösättningar


# Om du vill använda koden

Rapporten innehåller en del custom fields specifika för min båtklubb,
och har hårdkodade kolumnnamn. Det skulle gå att generalisera detta om nån är intresserad.

# Templates

Några filer skall ligga i templates-biblioteket. Det gäller
* en excelfil för schema-generering
* En varvskarta (pptx)
* Valfritt en fil med färgkoder.

## Färkgoder
Med filen `templates/colors.json` är det möjligt att definiera ett eget färgschema för kartan. Filen måste vara en giltig jsonfil, och skall ha följande format:

```json
{
    "reserved": [214, 245, 214],
    "declined": [255, 230, 230],
    "member_left": [255, 153, 255],
    "on_land": [230, 230, 255],
    "unknown": [255, 255, 255]
}
```

Nycklar utöver dessa kommer inte att beaktas.


# Filformat
## ex-members (`--exmembers`)

Default: `boatinfo/ex-members.txt`

Textfil där varje rad skall starta med ett medlemsnummer. En medlem per rad.
Exempel:
```
1   # Kalle kula har sålt båten
10 har inte kvar sin båt
# Följande medlemmar flyttade:
20
23
25
```

## sommarliggare (`--onland`)

Default: `boatinfo/sommarliggare.xlsx`

Excelfil med följande kolumner:
* År
* Medlemsnr

Övriga kolumner läses inte.

Syftet är att kunna markera de båtar som inte sjösatts denna sommar.

## Schemalagda (`--scheduled`)

Default: `boatinfo/torrsättning*.xlsx`

Excelfil med följande kolumner:
* Medlemsnr

Denna rapport kan med fördel genereras och laddas ner från BAS.

Övriga kolumner läses inte.

Syftet är att fånga de medlemmar som anmält torrsättning, men inte fyllt i formuläret.

## Medlemmar (`--members`)

Default: `boatinfo/Alla_medlemmar_inkl_båtinfo_*.xlsx`

Excelfil med följande kolumner:
* Medlemsnr
* Längd (båt)
* Bredd
* Förnamn
* Efternamn
* Plats

Denna rapport kan med fördel genereras och laddas ner från BAS.
