# Applied Genetics 1.2 (2007 - 2008)

Applied Genetics is a Windows desktop application for working directly with DNA and RNA sequences. It was written in Visual Basic 6.0 and brings together a sequence editor, restriction-enzyme analysis, codon translation, repeat detection, primer calculations, graphical annotation and a simple electrophoresis-gel simulation.

<hr>

![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/screen%20(VI).jpg)

<hr>

The repository preserves several stages of the project: the older Romanian Windows XP source, Romanian and English editions adapted for Windows 11, the original CD launcher, the file-association utility and the historical trial-key project. The Windows 11 editions keep the original interface, data files and project format while correcting a number of problems that appeared when the old VB6 code was compiled and used on a modern system.

<hr>

![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/DSC06704.JPG)

<hr>

Applied Genetics is mainly an interactive teaching and exploratory tool. It works with one nucleotide sequence at a time and shows the result directly in the editor, in a text output window or in one of the graphical views. Its algorithms are based largely on exact string matching. It is not intended to replace present-day sequence-alignment packages, genomic databases or clinical bioinformatics software.

## General operation

A sequence can be typed or pasted into the main editor, loaded from an Applied Genetics project or imported from a GenBank-style text file. While the sequence is being inspected, the program reports the current selection, its length, the number of `A`, `T`, `C`, `G` and `U` symbols, the nucleotide percentages and the `G+C` content. The editor is also used as the main annotation surface. Restriction sites, fragments, codons, repeats and other detected regions are marked by changing the foreground or background colour of the corresponding text. A horizontal map below the editor shows where the marked regions occur in relation to the complete sequence.

<hr>

<div align="left">
	<details>
	<summary>Applied Genetics 1.2 [+]</summary>
		<a href="">
		    <img src="https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/screen%20(I).jpg" alt="Applied Genetics 1.2">
		</a>
	</details>
</div>

<hr>


Characters outside the accepted nucleotide alphabet can be located and coloured in red. A cleaning command creates a version of the sequence containing only recognised nucleotide symbols. The importer follows the traditional GenBank layout: it looks for the sequence between `ORIGIN` and `//`, then removes line numbers, spaces and other formatting characters. The original code does not contain a dedicated FASTA parser.

The main areas of the program are summarised below.

| Area | Function |
|---|---|
| Sequence editor | Editing, selection statistics, colouring and sequence cleanup |
| Sequence conversion | Complement, reverse, DNA/RNA conversion and translation |
| Restriction analysis | Recognition-site detection, cleavage positions and fragment extraction |
| Gene tools | Reading-frame selection, codon translation and START/STOP marking |
| Inverse-sequence search | Reversed motifs, intervening regions and palindrome detection |
| Repetitive-sequence search | Exact tandem-repeat detection |
| Primer and probe tools | `G+C` content and melting-temperature estimates |
| Graphical output | Whole-sequence map, fragment views and simulated gels |
| Project system | Saving the sequence, output text and generated graphics |
| Report generator | Local HTML report with text and images |

## Sequence conversion and translation

The current selection can be complemented, reversed or converted between DNA and RNA. Complement and reverse are intentionally separate operations. The command named **Antisense** performs a nucleotide-by-nucleotide complement, but it does not reverse the sequence automatically.

A DNA selection can also be translated into an amino-acid sequence. DNA is first converted internally to RNA, after which the program reads the result in groups of three nucleotides. The reading-frame offset can be changed from the Gene panel, and the output can use either one-letter or three-letter amino-acid notation.

The codon table is editable. Each amino acid has its own list of RNA codons, so the translation rules can be inspected or modified from the Options window. START and STOP codons can be coloured in the editor and projected onto the graphical sequence map. In the historical implementation, STOP is represented internally as `Z` or as the text `[STOP]`.

A small sequence generator is included as well. It can create synthetic DNA or RNA strings of a requested length and place them in the OUTPUT window. One known defect remains in the historical DNA generator: it currently selects only `A`, `T` and `G`, because the random range contains three choices and therefore never reaches `C`.

## Restriction-enzyme analysis

Restriction enzymes are stored in the plain-text database `AG_DB.db`. The supplied database contains 126 records. Each line combines the enzyme name with its recognition sequence, using `[#]` as a separator and `*` to indicate the cleavage point.

```text
EnzymeName[#]RECOG*NITION
```

One or several enzymes can be selected for the same analysis. The program searches for exact recognition sequences in both displayed directions, records the sites, calculates the corresponding cut positions and combines the results into an ordered fragment map. Each fragment can then be extracted and described by length and nucleotide composition.


<hr>

<div align="center">

| Enzymes | Gene |
|---|---|
| ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/p1.png) | ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/p2.png) |


</div>

<hr>

The same result is presented in several forms. Recognition sites and fragment boundaries are coloured in the sequence editor, their positions are drawn on the whole-sequence map, and the calculated fragments are represented as bands in the gel view. The program can display a local gel for one enzyme or a general gel combining several enzymes.

The restriction panel also contains an experimental search for sites that could be formed after a deletion or insertion. The recognition sequence is divided at each possible internal position, and the two remaining parts are searched at a chosen separation distance. A single distance or an interval between 1 and 100 nucleotides can be tested. When a possible site is found, the left part, affected interval and right part are coloured separately. This is an exact textual search for potential sites rather than a general mutation-analysis engine.

## Inverse sequences, palindromes and direct repeats

The **I.S. / S.I.** module (Inverse Sequences / Secvențe Inverse) searches for a motif followed later by the same motif in reversed order. The program generates every possible motif for a selected alphabet and length interval, then scans the sequence for matching pairs. The normal alphabet is `A`, `C`, `G` and `T`, although it can be changed in the interface.

For every match, Applied Genetics records the position and length of the first motif, the sequence between the two motifs and the position of the reversed motif. The result is shown both as text and as a small graphical strip. When the distance between the two motifs is zero, the structure is identified as a palindrome in the generated report.

The search is exhaustive. Its running time rises quickly as the maximum motif length increases because the number of possible combinations grows exponentially. The interface shows the current combination, total number of combinations and processing speed, and the Windows 11 edition contains additional cancellation checks so that a long search can be stopped more reliably.

<hr>

<div align="center">

| Inverse Sequences | Repetitive Sequences |
|---|---|
| ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/p3.png) | ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/p4.png) |

</div>

<hr>

The **R.S. / S.R.** module (Repetitive Sequences / Secvențe Repetitive) uses a related generator to find exact tandem repeats. The user chooses the motif-length interval and the allowed number of consecutive repetitions. Matching regions are coloured in the editor and described in OUTPUT. The original algorithm does not allow mismatches, gaps or approximate repeat scoring.

## Gene and inverse-sequence details

The Gene panel also allows AUG codons to be delimited independently and lets the user define the text inserted at START and STOP positions. Ordinary codons can be displayed using alternating colours historically labelled as exon and intron colours.

For each inverse-sequence match, the program records the position and length of the first motif, the position and length of the intervening sequence, and the position of the reversed motif. The result includes both a textual representation and a graphical strip showing the two motifs and the region between them. During processing, the status area displays the current generated combination, the total number of combinations and the number of combinations tested per second.

## Primer and probe estimates

For a selected sequence, the program calculates the `G+C` percentage and estimates the melting temperature. Short selections are labelled as primers and longer selections as probes. The first formula is the familiar simple approximation:

```text
Tm = 4 × (G + C) + 2 × (A + T)
```

A second formula includes the sodium concentration:

```text
Tm = 81.5 + 16.6 × log10([Na+]) + 0.41 × (%G+C) − 675 / n
```

Here, `n` is the sequence length and `[Na+]` is expressed in mol/L. The constants used by both equations are editable from the Options window. These calculations reflect the methods implemented in the original program and are not nearest-neighbour thermodynamic calculations. Thus, `[Na+]` is the molar sodium-ion concentration. The default value used by the program is `0.05 mol/L`, corresponding to `50 mM Na+`.

## Integrated graphical representation

The central idea behind Applied Genetics was to connect the different structures detected in a nucleotide sequence within a single graphical environment. Restriction sites, fragments, codons, genes, repeats and inverse sequences are not presented as isolated results; they are linked to their exact positions in the original sequence and represented simultaneously in the text editor, the whole-sequence map and the corresponding graphical views. This connection between the sequence, its annotations and the generated diagrams makes it possible to move from a local result to its wider structural context. The purpose was to provide a clear and intuitive view of how the different elements relate to one another inside the same sequence, rather than displaying them as separate numerical or textual outputs.

## Graphical views and gel simulation

Applied Genetics produces a compact visual summary of most analyses. The whole-sequence map shows the location of sites, fragments, codons or repeat regions without requiring the entire nucleotide string to fit on screen. Longer result sets, especially inverse-sequence diagrams and individual enzyme gels, are displayed in a separate scrollable window.

The gel view converts calculated fragment lengths into horizontal bands. A time parameter can be changed to alter the apparent migration distance, and information about the enzyme, fragment number and fragment length is displayed while the result is inspected. The view is useful for explaining the relation between cleavage positions and fragment sizes, but it is a simplified graphical model rather than a calibrated simulation of a laboratory gel.

## Options and customization

The **Options** window brings together the main configuration settings of Applied Genetics. Its five tabs control the restriction-enzyme database, program colours, text appearance, codon assignments and primer calculations.

The **Restriction Enzymes** tab displays the complete enzyme list together with each recognition site. The supplied database contains 126 enzymes, and entries can be tested, added or removed directly from this window. The **Program Colors** tab defines the colours used throughout the graphical analyses, including START and STOP codons, genes, repetitive sequences, inverse sequences, normal restriction sites, insertion- and deletion-related sites, the vertical sequence marker, the graphical overview and the simulated gel.

<div align="center">

| Restriction Enzymes | Program Colors |
|---|---|
| ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/o1.png) | ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/o2.png) |

| Program Text | Codons |
|---|---|
| ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/o3.png) | ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/o4.png) |

| Primers | Visualisation |
|---|---|
| ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/o5.png) | ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/gif/sequence%20highlighting%20(II).gif) |

</div>


The **Program Text** tab controls the appearance of both the nucleotide editor and the OUTPUT window. Font family, size, foreground colour, background colour and styles such as bold, italic, underline and strikethrough can be changed independently. The **Codons** tab contains the RNA codon assignments used during translation and allows the mappings for each amino acid to be edited. Translation can be displayed using either one-letter notation, such as `M`, `P` and `Q`, or three-letter notation, such as `Met`, `Pro` and `Gln`.

The **Primers** tab contains the two melting-temperature formulas used by the program. Both the simple approximation and the sodium-adjusted formula can be selected, and their numerical constants can be modified directly from the interface.


## Applied Genetics project files

A saved project is made from three related files:

```text
ProjectName.pro
ProjectName_ag.agx
ProjectName_gl.agx
```

The `.pro` file contains the input sequence in RTF form, the contents of the OUTPUT window and several internal sections used when the project is reconstructed. The loader identifies those sections through the following markers:

```text
[#PROCESARE#]
[#PIC#]
[#NR_GEL#]
```

The `.agx` files are ordinary Windows bitmap data stored with application-specific extensions. `ProjectName_ag.agx` contains the graphical overview of the sequence, while `ProjectName_gl.agx` contains the general gel image. They are support files belonging to the `.pro` project and are not standard biological data formats.

A small helper program named `MS.exe` is associated with `.agx` files. Its purpose is to explain that these images are project resources rather than files intended to be opened and edited separately.

## HTML report generation

The **Compile** command creates a local report named `print.htm`. Depending on the analyses already performed, the report can include the project title, nucleotide sequence, OUTPUT text, sequence overview, combined gel, individual enzyme gels and inverse-sequence diagrams.

The graphical elements are exported as temporary JPEG files beside the report. The original preview routine tries to launch Internet Explorer from a fixed path. Report generation itself still works, but automatic preview should eventually be changed to open the file with the default browser.

## Repository contents

| Folder or archive | Description |
|---|---|
| `AG (EN Win 11 - 29.07.2026)` | English source adapted for Windows 11 |
| `AG (RO Win 11 - 29.07.2026)` | Romanian source adapted for Windows 11 |
| `AG (RO Win XP - 23.05.2007)` | Earlier Romanian Windows XP source |
| `AG Fisiere Asociate` | Source for the `.agx` file helper |
| `AG TRIAL KEY` | Historical activation-key utility and technical notes |
| `CD Interfata` | Source for the original CD launcher |

The two main Windows 11 source trees contain approximately 70 files each. Of these, 43 are VB6 forms, modules, classes, controls or property pages, amounting to roughly 38,600 lines of source code.

## Source structure

| File | Role in the project |
|---|---|
| `genetica.vbp` | Main Visual Basic 6 project |
| `genetica.frm` | Main window, project handling and report generation |
| `Mai_mult.frm` | Options, colours, codon editor and enzyme-database editor |
| `Functii_ENZIME.bas` | Restriction-site and insertion/deletion searches |
| `GEL_ELECTROFOREZA.bas` | Individual and combined gel rendering |
| `OUT.bas` | Translation, direct repeats, inverse sequences and formatted output |
| `Universal.bas` | General sequence operations and graphical helpers |
| `functii_ADN.bas` | DNA-to-RNA conversion |
| `Functii_ARN.bas` | RNA-to-DNA conversion |
| `Functii_PROTEINE.bas` | Protein-related routines |
| `coloreaza_inteligent.bas` | RichTextBox colouring and selection utilities |
| `clsBruteforce.cls` | Motif generator used by the repeat-search modules |
| `Lista_cu_geluri.frm` | Auxiliary window for gels and inverse-sequence graphics |
| `Generator.frm` | Synthetic nucleotide-sequence generator |
| `Trial.frm` and `Trial_COD.frm` | Historical trial and registration interface |
| `Kriptare.bas`, `kGen.bas`, `clsDS2.cls` | Registration-data processing |
| `Asociere_AG_fisier.bas` | Registration of `.pro` and `.agx` associations |
| `SaveBitmapAs.bas` | Image export through GDI+ |

The repository also includes the custom VB6 controls used by the interface and several classes for reading or writing BMP, GIF, ICO and PNG images.

## Runtime data

| File | Purpose |
|---|---|
| `AG_DB.db` | Restriction-enzyme database |
| `ajutor.htm` | Local help file |
| `scr.db` | HTML and JavaScript fragment inserted into reports |
| `sit.dat` | Optional web-address override |
| `aso.dat` | Marker showing that file associations were initialised |
| `ag_exp.001` | Stored registration information |
| `sys_exp.dll` | Trial-counter marker; it is not a real DLL |
| `MS.exe` | Helper associated with `.agx` files |
| `print.htm` | Generated HTML report |

## Building the application

The main project is `genetica.vbp` and must be opened in Microsoft Visual Basic 6.0. The project uses the VB6 runtime, Windows GDI+ and several standard 32-bit Microsoft controls.

| Component | Typical file |
|---|---|
| Visual Basic 6 runtime | `MSVBVM60.DLL` |
| Rich TextBox | `RICHTX32.OCX` |
| Common Dialog | `COMDLG32.OCX` |
| FlexGrid | `MSFLXGRD.OCX` |
| Tabbed Dialog | `TABCTL32.OCX` |
| Windows Common Controls | `MSCOMCTL.OCX` |

The custom controls used by Applied Genetics are already present in the source tree. Their binary companion files must remain beside the corresponding source files: `.frm` files require their `.frx` files, and `.ctl` files require their `.ctx` files.

After the required 32-bit controls have been installed and registered, either Windows 11 source tree can be opened in VB6. The startup object should remain `Trial`, and the main executable should be compiled as `AG.exe`. The database, help files, images and other runtime data must be copied beside the executable.

On 64-bit Windows, 32-bit OCX controls are normally registered with:

```text
C:\Windows\SysWOW64\regsvr32.exe
```

The auxiliary projects are separate from the main application. `Mesaj.vbp` builds the `.agx` helper, `Include.vbp` builds the CD interface, and `cheie.vbp` builds the historical key utility.

## Trial strategy

The original release used a 30-run trial system. Each time the program started, it read a usage counter from two separate locations: the current user’s VB settings area in the Windows Registry and a small local file named `sys_exp.dll`. Despite its extension, this file was not a real Windows library; it was only used as a secondary storage location for the trial counter.

The two values were compared at startup. If they matched, the counter was incremented and the program continued. If they differed, the application assumed that one of the values had been manually modified and treated the trial information as invalid. Keeping the same value in two different locations made simple counter resets more difficult, although the mechanism was still relatively easy to inspect and bypass by modern standards.

<div align="center">
	
| Trial Ro | Trial En |
|---|---|
| ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/trial%20RO.png) | ![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/trial%20EN.png) |

</div>

Registration was based on a username and a corresponding activation code. The activation window validated the entered pair and, when accepted, stored the registration information in an encrypted file named `ag_exp.001`. On later launches, the program checked this file before displaying the trial counter. A valid registration allowed the application to start without consuming one of the remaining trial runs.

The trial interface, activation-code generator, encryption routines, and related validation classes are preserved in the source code because they formed part of the original software architecture. They are useful for studying historical shareware protection techniques, but they should not be considered secure licensing methods for current software. The implementation relies on writable Registry entries, local files, reversible custom encryption, and Windows paths that may require elevated permissions on modern systems.

## Original CD distribution kit

The repository also preserves the presentation layer of the original CD edition. The `CD Interfata` directory contains the Visual Basic 6 project used for the disc launcher, while the `img` directory includes the surviving CD covers, disc artwork and other visual material associated with the physical release.




<hr>

<div align="center">
		<a href="">
		    <img src="https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/setup.png" alt="CD Applied Genetics 1.2">
		</a>
</div>

<hr>

![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/DVD%20(IV).png)

<hr>

The launcher was designed as the entry point of the distribution kit, providing access to the Applied Genetics installer and to a separate video presentation. In the original directory structure, it expected the installer at `Kit/AppliedGenetics_setup.exe` and the video interface at `video/index.htm`. The preserved launcher therefore documents how the software was intended to be presented and installed from the CD, even when some of the original installation or video files are not part of the current archive.


<hr>

![screenshot](https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/DVD%20(II).png)

<hr>

Together, the launcher source, interface resources, covers and disc artwork preserve more than the application itself: they reconstruct the appearance and organization of the original physical software package.


## Windows 11 edition

The modern source trees are not a rewrite. They retain the VB6 architecture and original project format, but repair several faults that became visible during compilation and use on Windows 11.

<hr>

<div align="center">
		<a href="">
		    <img src="https://github.com/Gagniuc/Applied-Genetics-1.2/blob/main/img/load.png" alt="CD Applied Genetics 1.2">
		</a>
</div>

<hr>

The form-name casing was made consistent with the internal VB6 attributes, parts of the main layout were widened, and protection was added against the re-entrant calls that could occur when `DoEvents` was used inside the brute-force engine. Long searches now respond better to the Stop command, including during result colouring. A repeat-search accumulator that was not reset correctly was fixed, and boundary checks were added around several `Mid$` operations near the beginning and end of a sequence.

The Windows 11 source is available in Romanian and English. Internal control names, data markers and project conventions were left unchanged wherever possible so that the translated edition would remain compatible with the original files.

## Historical trial system and file associations

The original distribution used a 30-run trial. The counter was stored once in the current user's VB settings in the Registry and once in a file named `sys_exp.dll`. The two values were compared in an attempt to detect manual changes. Successful registration data was encrypted and written to `ag_exp.001`.

This code is preserved because it belongs to the original application, not because it represents a current licensing design. It writes to old Registry and system-folder locations and may require administrator rights or further changes on a modern installation.

Applied Genetics can also register `.pro` files with `AG.exe` and `.agx` files with `MS.exe`. The association routine writes directly to `HKEY_CLASSES_ROOT` and uses `aso.dat` as a one-time initialisation marker.

## Technical limitations

| Limitation | Practical consequence |
|---|---|
| 32-bit VB6 and ActiveX dependencies | The program requires legacy runtime components |
| ANSI source and text handling | Unicode biological annotations are not handled reliably |
| Exact matching only | Mismatches, gaps and ambiguity-aware searches are absent |
| No native FASTA importer | FASTA sequences must be pasted or converted first |
| Exhaustive motif generation | Long inverse-repeat and repeat searches become very slow |
| Single-threaded interface | Intensive searches can temporarily reduce responsiveness |
| Fixed historical arrays | Very large sequences or result sets may exceed old limits |
| Approximate primer formulas | Values should be checked with modern primer software |
| Simplified gel rendering | Band positions are illustrative rather than experimentally calibrated |
| Internet Explorer preview | Automatic report opening needs modernisation |
| Internal `.agx` resources | Project graphics are not standard bioinformatics files |
| No automated test suite | Results must be checked against known examples when the code is changed |

VB6 binary companion files such as `.frx`, `.ctx`, `.pgx` and `.agx` should not be edited as ordinary text. The internal markers in `.pro` files must also remain unchanged because the loader searches for their exact spelling.

The repository is intended to preserve the original application in a form that can still be studied, compiled and used. Any future port to another language or framework should be kept separate from the historical VB6 edition so that the original program and its file formats remain available.

## Video demonstration

<details>
<summary>Video demonstration (I) [+]</summary>

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

</details>


<details>
<summary>Video demonstration (II) [+]</summary>

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

</details>
