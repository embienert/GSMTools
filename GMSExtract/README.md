# GMSExtract

### Description
This script is intended to be used for extracting the H-/P-/EUH-Statements, the WGK (Wassergefährdungsklasse) 
and CAS-Numbers from chemicals' safety data sheets. The extracted Statements will be returned
as TAB-separated Lists (can be modified in GMSExtract.py) of the COMMA-separated statements in order of their discovery
within the input. e.g.
```
aceton.pdf	67-64-1	H225, H319, H336	P210, P305 + P351 + P338, P403 + P233	EUH066	1
```

Input may be the path to a pdf file, finished with an empty input line. `*.pdf` will read all PDFs in the current directory. <br>
If a path to a pdf file was recognized the program will confirm, that the input will be interpreted as such.
```
> aceton.pdf
>

Interpreting input as path of pdf file.
```

Input may alternatively be the entire or partial safety data sheet as text, finished with an empty input line.
Placement of linebreaks is irrelevant, as they will be replaced with whitespaces in the process. This allows
text to be copy/pasted into the command line.
```
> Sicherheitsinformationen gemäß GHS Gefahrensymbol(e)
> Gefahrenhinweis(e)
> H315: Verursacht Hautreizungen.
> H319: Verursacht schwere Augenreizung.
> H335: Kann die Atemwege reizen.
> Sicherheitshinweis(e)
> P261: Einatmen von Staub/ Rauch/ Gas/ Nebel/ Dampf/ Aerosol vermeiden.
> P271: Nur im Freien oder in gut belüfteten Räumen verwenden.
> P280: Schutzhandschuhe/ Augenschutz/ Gesichtsschutz tragen.
> P302 + P352: BEI BERÜHRUNG MIT DER HAUT: Mit viel Wasser waschen.
> P305 + P351 + P338: BEI KONTAKT MIT DEN AUGEN: Einige Minuten lang behutsam mit Wasser spülen.
> Eventuell vorhandene Kontaktlinsen nach Möglichkeit entfernen. Weiter spülen.
> Ergänzende Gefahrenhinweise EUH061
> SignalwortAchtungLagerklasse10 - 13 Sonstige Flüssigkeiten und
> FeststoffeWGKWGK 1 schwach wassergefährdendEntsorgung3
>
```

### Versions

Python: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 3.8.5 <br>
pdfminer.six: &nbsp;&nbsp; 20201018
