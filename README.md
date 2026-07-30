# Applied Genetics 1.2 (2007 - 2008)

**A historical Visual Basic 6 desktop environment for interactive DNA/RNA sequence analysis, restriction mapping, motif discovery, codon translation, primer evaluation, and simulated gel electrophoresis.**

> [!IMPORTANT]
> Applied Genetics is preserved as historical, educational, and research software. It is not a clinical or diagnostic application, and its results should be independently verified with current bioinformatics tools.

## Overview

**Applied Genetics 1.2** is a native Windows application developed for the interactive examination of nucleotide sequences. It combines sequence editing, exact motif detection, restriction-enzyme analysis, codon translation, direct- and inverse-repeat discovery, primer melting-temperature estimation, graphical sequence annotation, and a simplified agarose-gel simulation in a single desktop interface.


![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/load.png)


The original program was created in the academic environment of the **University of Bucharest, Institute of Genetics**, during 2008–2009. The source was later recovered and adapted for modern 32-bit VB6 execution under Windows 11, while preserving the original interface, algorithms, project format, terminology, and historical shareware subsystem.

![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/DVD%20(IV).png)

<hr>

![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/DVD%20(II).png)

<hr>

![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/DSC06704.JPG)

The application is primarily a **DNA/RNA analysis tool with derived protein output**. It does not perform sequence alignment, database similarity searches, phylogenetic inference, or modern high-throughput analysis.

## Project facts

| Item | Description |
|---|---|
| Product | Applied Genetics |
| Version | 1.2 beta |
| Original period | 2007–2008 |
| Restoration | July 2026 |
| Language | Microsoft Visual Basic 6.0 |
| Platform | 32-bit Windows desktop |
| Interface languages | English and Romanian |
| Main executable name | `AG.exe` |
| Main project | `genetica.vbp` |
| Main form | `genetica.frm` / `Principal` |
| Restriction-enzyme database | 126 enzyme records |
| Historical distribution model | Shareware, 30-use trial |
| Current purpose | Preservation, education, code archaeology, and historical research |

Each principal source tree contains approximately **70 files**, including **43 VB6 source files** and roughly **38,600 lines** across forms, modules, classes, user controls, and the property page.

## Main capabilities

### 1. Interactive sequence workspace

The main window is built around a large RichTextBox used for nucleotide input and annotation. A sequence can be entered manually, pasted, loaded from an Applied Genetics project, or imported from a GenBank-style flat file.

![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/screen%20(I).jpg)

<hr>

![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/gif/screen%20(I).gif)




The workspace provides:

- editable DNA or RNA sequence text;
- selection start, selection end, and selected-length indicators;
- real-time counts for `A`, `T`, `C`, `G`, and `U`;
- percentage composition for the current selection;
- graphical projection of selected and detected regions across the complete sequence;
- configurable foreground, background, font, size, and text style;
- manual sequence highlighting;
- detection and red highlighting of characters outside `A/T/C/G/U`;
- generation of a cleaned sequence containing only recognized nucleotide symbols;
- optional accumulation of new results in the OUTPUT window instead of replacing previous results.

The GenBank importer searches for an `ORIGIN` section terminated by `//`, removes numbering and formatting characters, and retains recognized nucleotide symbols. Native FASTA parsing is not implemented.

### 2. Sequence transformations

Applied Genetics can transform the current selection and send the result to OUTPUT:

- **Antisense/complement** — replaces each DNA nucleotide with its complementary base;
- **Reverse** — reverses the order of the selected sequence;
- **RNA → DNA** — replaces uracil with thymine;
- **DNA → RNA** — replaces thymine with uracil;
- **DNA → amino acids** — converts DNA to RNA codons and translates them using the active codon table.

The historical interface treats complement and reverse as separate operations. The command labeled “Antisense” performs a base-wise complement; it does not automatically combine complementing with reversal.

A small generator can also create synthetic DNA or RNA strings of a requested length and place them in OUTPUT.

### 3. Restriction-enzyme analysis

The restriction-analysis engine uses the plain-text database `AG_DB.db`, which contains **126 named enzymes**. Each record has the form:

```text
EnzymeName[#]RECOG*NITION
```

The `*` character marks the cleavage position inside the recognition sequence.

For one or more selected enzymes, the program can:

- detect exact recognition sites in the displayed 5′→3′ sequence;
- detect the corresponding complementary-direction sites labeled 3′→5′;
- count sites separately for both directions;
- merge detected cut positions into a common ordered fragment map;
- extract the resulting restriction fragments;
- report fragment length and nucleotide composition;
- color recognition sites and fragment boundaries in the sequence editor;
- project restriction positions onto the graphical sequence overview;
- render fragment bands on the simulated gel.

#### In-silico deletion and insertion searches

In addition to normal restriction mapping, the application searches for **potential recognition sites that could be formed by a deletion or insertion**.

The user can select:

- deletion in either direction;
- insertion in either direction;
- one exact distance or a distance interval;
- distances from 1 to 100 nucleotides.

The algorithm splits the recognition sequence at possible internal positions and checks whether its two parts occur at the required separation. Matching regions are highlighted with independently configurable colors for the left segment, affected interval, and right segment.

This is an exact string-based exploration of possible sites; it is not a general mutation simulator or sequence-variant annotation system.

### 4. Gene and codon tools

The Gene tab translates the nucleotide sequence codon by codon and provides:

- adjustable reading-frame offset;
- DNA-to-RNA conversion before translation;
- configurable amino-acid notation;
- highlighting of START and STOP codons;
- optional delimitation of `AUG`;
- customizable text inserted at START and STOP positions;
- graphical marking of codon events;
- alternating historical “exon/intron” display colors for ordinary codons.

The default codon editor contains the standard RNA codon assignments for 20 amino acids plus STOP. The user can redefine the codon list associated with each amino acid.

Two output styles are available:

- short notation: `M`, `P`, `Q`, and so on;
- long notation: `Met`, `Pro`, `Gln`, and so on.

In the historical implementation, STOP is represented internally as `Z` or `[STOP]`.

### 5. Inverse sequences and palindromic structures

The **I.S. / S.I.** module searches for inverse sequence pairs using a configurable brute-force generator.

The user can define:

- the character set, normally `A`, `C`, `G`, and `T`;
- minimum motif length;
- maximum motif length;
- an optional starting combination;
- whether graphical results should be displayed.

For every generated motif, the program checks whether the motif occurs in the sequence and whether its reversed form occurs later. Matches are reported with:

- motif position;
- motif length;
- intervening-sequence position and length;
- second motif position;
- textual representation of both motifs and their spacer;
- a graphical strip showing the two motifs and the sequence between them.

A pair with a zero-length spacer is labeled as a palindrome in the generated HTML report. The live status area displays the current combination, total combinations, and combinations per second.

Because the search enumerates all possible strings in the selected length interval, its complexity grows exponentially with alphabet size and motif length. Practical searches should therefore use short motifs and narrow ranges.

### 6. Direct repetitive-sequence discovery

The **R.S. / S.R.** module uses the same brute-force motif generator to locate direct tandem repeats.

The user specifies:

- minimum motif length;
- maximum motif length;
- minimum number of repetitions;
- maximum number of repetitions.

When a repeated motif is detected, the corresponding sequence region is highlighted and a description is appended to OUTPUT. The implementation is designed for exact contiguous repeats and does not support mismatches, gaps, approximate tandem repeats, or statistical scoring.

### 7. Primer and probe statistics

For short selected regions, the interface displays:

- selected length;
- `G+C` percentage;
- an estimated melting temperature;
- the historical label “Primer” for short selections and “Probe” for longer selections.

Two editable formulas are included:

#### Approximate formula

```text
Tm = 4 × (G + C) + 2 × (A + T)
```

#### Salt-adjusted formula

```text
Tm = 81.5 + 16.6 × log10([Na+]) + 0.41 × (%G+C) − 675 / n
```

where `n` is the sequence length and `[Na+]` is the molar sodium concentration.

The constants used by both formulas can be modified from the Options window. These are historical approximations, not nearest-neighbor thermodynamic calculations.

### 8. Sequence and gel visualization

Applied Genetics uses several coordinated visual layers:

- color formatting directly inside the nucleotide RichTextBox;
- a horizontal whole-sequence map;
- algorithm-specific colored segments;
- a local gel panel for the active enzyme;
- a general gel combining results from several enzymes;
- an auxiliary scrollable window for inverse-sequence diagrams;
- live information showing enzyme, fragment number, fragment length, and migration distance.

Gel migration is a simplified visual model based primarily on fragment length and a user-adjustable time parameter. It is intended as an explanatory simulation rather than a physical electrophoresis model.

### 9. Project persistence

Applied Genetics saves a project as three coordinated files:

```text
ProjectName.pro
ProjectName_ag.agx
ProjectName_gl.agx
```

The `.pro` file stores:

- the RTF-formatted input sequence;
- the text from OUTPUT;
- internal section markers;
- the number of generated gel controls.

The main internal markers are:

```text
[#PROCESARE#]
[#PIC#]
[#NR_GEL#]
```

The two `.agx` files contain standard Windows bitmap data under a project-specific extension:

- `_ag.agx` — the graphical whole-sequence representation;
- `_gl.agx` — the general gel image.

A sample project is included as `Proiect1.pro`, `Proiect1_ag.agx`, and `Proiect1_gl.agx`.

### 10. HTML report generation

The toolbar command labeled **Compile** builds a local HTML report containing:

- project title;
- graphical sequence overview;
- general gel image;
- individual gel images;
- inverse-sequence diagrams;
- nucleotide sequence formatted in fixed-width lines;
- complete OUTPUT text.

The report generator creates `print.htm` and temporary JPEG files. The original code then attempts to open the report through a hard-coded Internet Explorer path. The HTML generation itself remains useful, but automatic preview should be updated to use the system’s default browser on current Windows versions.

## Typical workflow

1. Start the application and continue through the historical trial screen.
2. Paste a sequence or import a GenBank flat file.
3. Run the invalid-character check and inspect nucleotide composition.
4. Choose one of the four analysis tabs:
   - restriction enzymes;
   - gene/codon processing;
   - inverse sequences;
   - repetitive sequences.
5. Configure the relevant lengths, directions, colors, formulas, or enzyme set.
6. Press **Start** or **Compile**, depending on the selected operation.
7. Inspect the colored input sequence, OUTPUT text, graphical map, and gel.
8. Save the work as a `.pro` project with its two `.agx` resources.
9. Generate an HTML report when a portable visual summary is needed.

## File formats and runtime data

| File | Purpose |
|---|---|
| `genetica.vbp` | Main VB6 project |
| `AG_DB.db` | Plain-text restriction-enzyme database |
| `*.pro` | Applied Genetics project: RTF input plus textual result sections |
| `*_ag.agx` | Whole-sequence map stored as a BMP |
| `*_gl.agx` | General gel stored as a BMP |
| `aso.dat` | Marker showing that file associations were initialized |
| `scr.db` | HTML/JavaScript fragment inserted into compiled reports |
| `ajutor.htm` | Local help and project documentation |
| `sit.dat` | Optional override for the project’s online URL |
| `ag_exp.001` | Encrypted historical registration data |
| `sys_exp.dll` | Historical trial-counter marker; not an actual DLL |
| `MS.exe` | Helper opened when a user double-clicks an `.agx` resource |
| `print.htm` | Generated HTML report |
| `gel.jpg`, `graf.jpg`, `SI*.jpg`, `sub_gel*.jpg` | Generated report images |

## Source-code architecture

| Component | Role |
|---|---|
| `genetica.frm` | Main user interface, event coordination, project I/O, live selection statistics, report generation |
| `Mai_mult.frm` | Options, colors, fonts, primer formulas, codon editor, enzyme database editor |
| `OUT.bas` | Direct repeats, inverse sequences, codon translation, formatted output |
| `Functii_ENZIME.bas` | Restriction-site detection and insertion/deletion site exploration |
| `GEL_ELECTROFOREZA.bas` | Local and general gel rendering |
| `Universal.bas` | Sequence transformations, random generation, composition statistics, graphical helpers |
| `functii_ADN.bas` | DNA-to-RNA conversion |
| `Functii_ARN.bas` | RNA-to-DNA conversion |
| `Functii_PROTEINE.bas` | Historical protein-related routines |
| `coloreaza_inteligent.bas` | RichTextBox highlighting and selection-color utilities |
| `clsBruteforce.cls` | Stateful motif-combination generator and performance statistics |
| `Trial.frm`, `Trial_COD.frm` | Historical trial and activation user interface |
| `clsDS2.cls`, `kGen.bas`, `Kriptare.bas` | Historical registration encryption and key validation |
| `Asociere_AG_fisier.bas` | Windows Registry association for `.pro` and `.agx` |
| `Special_DIR.bas` | Windows special-folder resolution |
| `SaveBitmapAs.bas` | GDI+ image export |
| `cBMPparser.cls`, `cGIFparser.cls`, `cICOparser.cls`, `cPNGparser.cls`, `cPNGwriter.cls` | Image parsing and writing support |
| `aicAlphaImage.ctl`, `isButton.ctl`, `jcFrames.ctl`, `LCD.ctl`, `zulu.ctl` | Custom visual controls |
| `Lista_cu_geluri.frm` | Auxiliary window for gels and inverse-sequence graphics |
| `Generator.frm` | Synthetic nucleotide-sequence generator |
| `Actiuni.frm` | Search/navigation actions |
| `Despre.frm` | About dialog |
| `Intro.frm` | Splash/loading form |

## Included archives and variants

| Archive or folder | Contents |
|---|---|
| `AG (EN Win 11 - 29.07.2026)` | English source tree adapted for Windows 11 |
| `AG (RO Win 11 - 29.07.2026)` | Romanian source tree adapted for Windows 11 |
| `AG (RO Win XP - 23.05.2007)` | Legacy Romanian Windows XP-oriented source snapshot |
| `AG Fisiere Asociate` | Source for the `.agx` resource-message helper compiled as `MS.exe` |
| `AG TRIAL KEY` | Historical registration-key utility and technical notes |
| `CD Interfata` | Source for the original CD launcher and installer-copy interface |

The main archives contain source code and resources. A compiled `AG.exe` is not present in the analyzed source packages and must be built from `genetica.vbp` unless a separate release binary is provided.

## Building from source

### Requirements

- Microsoft Visual Basic 6.0, preferably with Service Pack 6;
- the 32-bit VB6 runtime (`MSVBVM60.DLL`);
- Microsoft Rich TextBox Control: `RICHTX32.OCX`;
- Microsoft Common Dialog Control: `COMDLG32.OCX`;
- Microsoft FlexGrid Control: `MSFLXGRD.OCX`;
- Microsoft Tabbed Dialog Control: `TABCTL32.OCX`;
- Microsoft Windows Common Controls: `MSCOMCTL.OCX`;
- Windows GDI+;
- a 32-bit ActiveX-compatible environment.

The custom controls and parser classes used by Applied Genetics are included as source. The Microsoft OCX dependencies are not embedded in the project and must be installed from legitimate Microsoft/VB6 media or redistributables.

On 64-bit Windows, 32-bit OCX controls are normally registered using the 32-bit registration tool located in `C:\Windows\SysWOW64\regsvr32.exe`.

### Compilation steps

1. Select the English or Romanian Windows 11 source directory.
2. Keep every `.frm` beside its `.frx` file and every `.ctl` beside its `.ctx` file.
3. Ensure all required Microsoft OCX controls are installed and registered.
4. Open `genetica.vbp` in the VB6 IDE.
5. Resolve any missing references without changing the control versions unless necessary.
6. Confirm that the startup object is `Trial`.
7. Compile the project as `AG.exe`.
8. Place the required runtime data and image files beside `AG.exe`.
9. Keep `MS.exe` beside the main executable if `.agx` file association is required.

Optional auxiliary builds:

- open `Mesaj.vbp` to rebuild `MS.exe`;
- open `Include.vbp` to rebuild the historical CD interface;
- open `cheie.vbp` to inspect or rebuild the historical registration utility.

## Windows 11 restoration

The 2026 source trees retain the original VB6 design while repairing several failures encountered in modern environments:

- form-name casing conflicts between `Begin VB.Form` and `Attribute VB_Name` were corrected;
- the main layout was widened for modern display behavior;
- the brute-force engine gained protection against re-entrant calls caused by `DoEvents`;
- the long-running inverse- and repeat-search loops now process Stop requests more reliably;
- internal scan routines check for cancellation while highlighting matches;
- the repetitive-search accumulator is reset between motif tests, preventing progressive slowdown;
- invalid `Mid$` start and length values near sequence boundaries are clamped;
- the project’s online address was updated to the GitHub repository;
- a fully translated English interface was added while preserving internal identifiers and file markers.

## Historical trial and activation subsystem

The original software was distributed as shareware with a **30-use trial counter**.

The preserved subsystem:

- stores a counter in the current user’s VB/VBA settings area in the Registry;
- mirrors the counter in a file named `sys_exp.dll` inside the Windows system-folder path;
- compares both values to detect manual modification;
- accepts a username and matching registration code;
- stores validated registration data in encrypted form in `ag_exp.001`;
- starts the main application through the `Trial` and `Intro` forms.

This mechanism is included for historical completeness. It should not be considered secure licensing by modern standards, and it writes to locations that may require elevated permissions on current Windows systems.

## Windows file associations

At first launch, Applied Genetics can register:

```text
.pro  → AG.exe
.agx  → MS.exe
```

The association code writes directly to `HKEY_CLASSES_ROOT` and uses `aso.dat` as a one-time marker.

Opening a `.pro` file passes its path to `AG.exe`, which loads the project. Opening an `.agx` file starts the small `MS.exe` helper, which identifies it as a resource belonging to an Applied Genetics project instead of attempting to edit it independently.

Direct Registry modification may require elevated permissions and may be blocked by modern Windows security policy.

## Known limitations

- The application is a historical 32-bit VB6 program dependent on legacy ActiveX controls.
- Source files are ANSI/Windows-1252 rather than Unicode.
- Sequence analysis is based on exact substring matching; ambiguity symbols and approximate matching are not supported.
- Native FASTA import is absent; the importer expects a GenBank-style `ORIGIN ... //` section.
- The inverse- and repeat-search modules use exhaustive brute force and become impractical as motif length grows.
- The interface is single-threaded and uses `DoEvents` to remain partially responsive during long calculations.
- Several internal arrays have fixed historical limits, making very large datasets unsuitable.
- Primer temperatures use simple formulas and are not replacements for modern nearest-neighbor calculations.
- The gel is a visual simulation, not a calibrated physical model.
- HTML report preview still targets Internet Explorer through a hard-coded path.
- The current DNA random-generator implementation selects only `A`, `T`, and `G`; cytosine is omitted because the random bound uses three choices instead of four.
- Project resources use custom extensions and are not interoperable with standard bioinformatics software.
- The original CD launcher expects `Kit/AppliedGenetics_setup.exe` and `video/index.htm`; these assets are not present in the analyzed `CD Interfata` archive.
- No automated test suite or formal scientific-validation dataset is included.

## Preservation and contribution notes

When modifying the historical source:

- preserve Windows-style CRLF line endings and the original ANSI encoding where possible;
- never separate `.frm` files from their `.frx` companions;
- never separate `.ctl` files from their `.ctx` companions;
- do not manually edit binary `.frx`, `.ctx`, `.pgx`, or `.agx` files;
- preserve the internal project markers exactly;
- keep form-name casing consistent with `Attribute VB_Name`;
- test both Romanian and English source trees after shared algorithm changes;
- retain the original source in an archival branch before major modernization.

Useful modernization tasks include:

- replacing Internet Explorer launch calls with the default system browser;
- adding FASTA import and export;
- replacing brute-force motif discovery with indexed or linear-time algorithms;
- adding IUPAC ambiguity-code support;
- removing or isolating the historical trial subsystem in archival builds;
- adding automated tests with known sequences and expected results;
- exporting results as CSV, JSON, SVG, or standard biological formats;
- porting the interface to a maintained platform while preserving the original VB6 edition.

## Scientific-use notice

Applied Genetics illustrates the design of an early desktop bioinformatics environment and can be valuable for teaching algorithmic ideas, inspecting historical software engineering, or reproducing the original interface.

It should not be used as the sole basis for:

- clinical interpretation;
- diagnostic decisions;
- primer ordering;
- restriction-digest planning;
- publication-grade sequence annotation;
- safety-critical laboratory work.

All results should be checked with current, validated software and appropriate experimental controls.

## Suggested citation

```text
Gagniuc, P. A. Applied Genetics 1.2 [Computer software].
University of Bucharest, Institute of Genetics, 2008–2009.
Source restoration and Windows 11 adaptation, 2026.
```

## Credits

**Original author:** Paul A. Gagniuc  
**Academic coordinator:** Prof. Dr. Lucian Gavrilă  
**Original institutional context:** University of Bucharest, Institute of Genetics  
**Copyright notice in the original project:** © 2008–2009

## License

No explicit software license is included in the analyzed historical archives. Until a repository-level `LICENSE` file is added, the source should be treated as **all rights reserved**.

---

Applied Genetics 1.2 is preserved as a record of early Windows-based bioinformatics software: a complete, highly visual VB6 environment in which sequence editing, genetic-code processing, restriction analysis, motif discovery, and gel simulation were integrated into one application.

## Video demonstration (I)

<p align="center">
  <a href="https://www.youtube.com/watch?v=nmyNR0VMKF0">
    <img src="https://img.youtube.com/vi/nmyNR0VMKF0/maxresdefault.jpg"
         alt="Applied Genetics 1.2 — Video Demonstration"
         width="800">
  </a>
</p>

<p align="center">
  <a href="https://www.youtube.com/watch?v=nmyNR0VMKF0">
    ▶ Watch the Applied Genetics 1.2 demonstration on YouTube
  </a>
</p>


## Video demonstration (II)

<p align="center">
  <a href="https://www.youtube.com/watch?v=CqIFUx_PaPA">
    <img src="https://img.youtube.com/vi/CqIFUx_PaPA/maxresdefault.jpg"
         alt="Applied Genetics 1.2 — Video Demonstration"
         width="800">
  </a>
</p>

<p align="center">
  <a href="https://www.youtube.com/watch?v=CqIFUx_PaPA">
    ▶ Watch the Applied Genetics 1.2 demonstration on YouTube
  </a>
</p>




