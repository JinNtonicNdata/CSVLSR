# Dashboard of a living review of interventions to prevent childhood sexual violence victimisation and perpetration

R Shiny code for an interactive dashboard presenting up-to-date evidence
on evaluations of interventions, laws and policies to prevent childhood
sexual violence (CSV) victimisation and perpetration

The dashboard is developed by the [Safe Futures
Hub](https://www.safefutureshub.org/) and the [Department of Social
Policy and Intervention, University of
Oxford](https://www.spi.ox.ac.uk/home).

-   **Link to the dashboard:**
    <https://safefutureshub.shinyapps.io/CSVlivingsystematicreview/>
-   **Review protocol / materials:** <https://osf.io/d2srw/overview>

> **This repository shares the code only.** The evidence dataset, fonts
> and logos are **not** included (see [What is not
> included](#what-is-not-included)). The app will not run end to end
> until you supply your own data in the expected format.

## What the dashboard does

The app (`app.R`) has two main tabs:

| Tab | Description |
|----|----|
| **Prevention explorer** | Only includes interventions that found a reduction in at least on CSV victimisation and perpetration outcome and no increase in any CSV victimisation and perpetration outcomes, with an interactive table. |
| **Studies included** | The full reviewed evidence base, with summary figures, an interactive world map, and interactive tables. |

## Repository contents

```         
app.R          # The whole Shiny app (UI + server), organised in 12 numbered sections
renv.lock      # Pinned R package versions (managed by renv)
renv/          # renv activation files
.Rprofile      # Activates renv when the project is opened
CSV LSR.Rproj  # RStudio project file
```

`app.R` is organised into 12 numbered sections: (1) packages and
options, (2) brand colours and theme, (3) data loading, (4) data
cleaning helpers, (5) derived list columns and dropdown choices, (6)
World Bank income/region join, (7) evidence-rating and badge helpers,
(8) derived analytical columns and usage logging, (9) country-map
helpers, (10) UI definition, (11) server definition and (12) app launch.

## Requirements

-   R (\>= 4.2 recommended)
-   The packages listed in `renv.lock`. The main ones are: `shiny`,
    `shinythemes`, `leaflet`, `sf`, `tmap`, `dplyr`, `tidyr`, `tibble`,
    `readxl`, `openxlsx`, `stringr`, `lubridate`, `ggplot2`, `DT`,
    `shinycssloaders`, `htmltools`, `knitr`, `sysfonts`, `showtext`.

## Getting started

1.  Clone the repository and open `CSV LSR.Rproj` in RStudio.

2.  Restore the package versions:

    ``` r
    renv::restore()
    ```

3.  Add the required input files (see below) to the project folder.

4.  Run the app:

    ``` r
    shiny::runApp()
    ```

## What is not included {#what-is-not-included}

For data governance and licensing reasons, these files are deliberately
**not** in the repository. `app.R` expects them in the app directory:

| File | Purpose |
|----|----|
| `CSVLSR_QA_subset.xlsx` | Extracted evidence data (sheet `Sheet 1`). |
| `World Bank Classification.xlsx` | Country income-group / region lookup. |
| `www/*.otf`, `www/*.ttf` | Brand fonts (GT Pressura Extended, General Sans, DM Mono). |
| `www/*.png`, `www/*.jpg` | Safe Futures Hub and DSPI logos and images. |

To adapt the code to your own evidence base, supply a spreadsheet with
the same column structure as the original extraction and update the file
paths in section 3 (*Data loading*) of `app.R`.

## Deployment

The live app is deployed to [shinyapps.io](https://www.shinyapps.io/)
with `rsconnect`. Deployment configuration (`rsconnect/`) is not tracked
in this repository.

## Citation

Department of Social Policy and Intervention and Safe Futures Hub. 2026.
An interactive dashboard on interventions to prevent child sexual
violence. Available from:
<https://safefutureshub.shinyapps.io/CSVlivingsystematicreview/>

## Contact and feedback

Feedback and suggestions for additional evidence can be submitted
through the forms in the dashboard.

For questions about the materials in this repository, contact Jinseo Kim
(<jinseo.kim@gtc.ox.ac.uk>). 🌱

## Licence

This work is licensed under the Creative Commons Attribution 4.0
International Public License (CC BY 4.0). See
<https://creativecommons.org/licenses/by/4.0/> for the full licence text.
