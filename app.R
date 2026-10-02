# =============================================================================
# Interactive dashboard on interventions to prevent childhood sexual violence
# -----------------------------------------------------------------------------
# A Shiny application presenting systematically reviewed global evidence on
# interventions and policies to prevent childhood sexual violence (CSV).
#
# Tabs:
#   - Prevention explorer : effective interventions only, summarised to
#                             intervention level.
#   - Studies included      : the full study-level evidence base, with summary
#                             charts and an interactive world map.
#
# Expected files in the app directory:
#   - CSVLSR_QA_subset.xlsx          (extracted evidence data, sheet "Sheet 1")
#   - World Bank Classification.xlsx (country income/region lookup)
#   - www/GT-Pressura-Extended-Regular.otf
#   - www/GeneralSans-Regular.otf
#   - www/DMMono-Regular.ttf
#   - www/sfh_full_colour_centralised_rgb.png
#   - www/DSPI_Primary_logo.jpg
#   - www/DSPI_Primary_logo_Ox.jpg
#
# File structure:
#   1.  Packages and global options
#   2.  Brand colours, fonts and ggplot theme
#   3.  Data loading
#   4.  Data standardisation and cleaning helpers
#   5.  Derived list columns and dropdown choices
#   6.  World Bank income / region join
#   7.  Coercions, evidence-rating helpers and badge helpers
#   8.  Derived analytical columns and event logging
#   9.  Country-map helper functions
#   10. UI definition
#   11. Server definition
#   12. App launch
# =============================================================================


# =============================================================================
# 1. Packages and global options
# =============================================================================

library(shiny)
library(shinythemes)
library(leaflet)
library(sf)
library(dplyr)
library(readxl)
library(stringr)
library(DT)
library(shinycssloaders)
library(lubridate)
library(ggplot2)
library(tidyr)
library(tibble)
library(htmltools)
library(knitr)
library(openxlsx)   # Excel downloads
library(sysfonts)
library(showtext)

sf::sf_use_s2(FALSE)

tags <- shiny::tags


# =============================================================================
# 2. Brand colours, fonts and ggplot theme
# =============================================================================

# --- Fonts -------------------------------------------------------------------
# The font files are not distributed with the code. If they are missing, the
# app still runs and R plots fall back to the default system fonts.
font_files <- c(
  "GT Pressura Extended" = "www/GT-Pressura-Extended-Regular.otf",
  "General Sans"         = "www/GeneralSans-Regular.otf",
  "DM Mono"              = "www/DMMono-Regular.ttf"
)
for (fam in names(font_files)) {
  if (file.exists(font_files[[fam]])) font_add(fam, font_files[[fam]])
}
showtext_auto()

# --- Safe Futures Hub brand palette (Brand Style Guide, Palette pp. 20-21) ---
# Primary
sfh_green    <- "#4F9B90"   # Community Green
sfh_orange   <- "#F09E56"   # Sunrise Orange
# Secondary
sfh_blue     <- "#5570CA"   # Horizon Blue
sfh_ocean    <- "#355571"   # Deep Ocean
# Tertiary
sfh_purple   <- "#8B5C8B"   # Purple
sfh_pink     <- "#F17592"   # Pink
# Mono
sfh_charcoal <- "#313131"   # Charcoal
sfh_sand     <- "#F3F2E8"   # Sand
# Lighter tones of Horizon Blue (Tones, p. 21), used for the map colour scale
sfh_blue_mid   <- "#99A9DF"
sfh_blue_light <- "#DDE2F4"
# Mono tones (Tones, p. 21): secondary text and light rules
sfh_charcoal_mid <- "#5A5A5A"
sfh_sand_dark    <- "#C2C2BA"

# --- Plot colour shortcuts ---------------------------------------------------
bg_white    <- "#FFFFFF"        # page background
orange_main <- sfh_orange       # bars
teal_main   <- sfh_green        # bars
dark_purple <- sfh_purple       # bars

# --- ggplot theme ------------------------------------------------------------
theme_sfh <- function(base_size = 13, base_family = "General Sans") {
  theme_minimal(base_size = base_size, base_family = base_family) +
    theme(
      plot.title         = element_text(family = "GT Pressura Extended",
                                        face = "bold", colour = sfh_charcoal),
      axis.title         = element_text(family = base_family, colour = sfh_charcoal),
      axis.text          = element_text(family = base_family, colour = sfh_charcoal),
      panel.grid.major.y = element_blank(),
      panel.grid.minor   = element_blank(),
      legend.position    = "none",
      plot.background    = element_rect(fill = "white", colour = NA),
      panel.background   = element_rect(fill = "white", colour = NA)
    )
}

# Apply the SFH theme as the default for all ggplots
theme_set(theme_sfh())


# =============================================================================
# 3. Data loading
# =============================================================================
# The dataset can be supplied in the global environment as `data_ext1`,
# otherwise it is read from the Excel file below.

if (exists("data_ext1")) {
  df_raw <- data_ext1
} else {
  path <- "CSVLSR_QA_subset.xlsx"
  df_raw <- tryCatch(
    readxl::read_excel(path, sheet = "Sheet 1"),
    error = function(e) {
      stop("Failed to read Excel file at ", path, " (", conditionMessage(e),
           "). Provide data_ext1 or correct path; see the README for the ",
           "expected input files.")
    }
  )
}

# Keep the Excel file's original header text before it gets mangled into
# syntactically-valid R names below, so downloads can show people a normal-
# looking column name instead of the dot-separated working name.
df_raw_original_colnames <- colnames(df_raw)

# Make column names syntactically valid and unique
colnames(df_raw) <- make.unique(make.names(colnames(df_raw)))

# Mangled name -> original Excel header text (e.g. "Study.Population" ->
# "Study Population"). A handful of the Excel headers already contained
# dots/underscores themselves (e.g. "Intervention.ID"), so those still need
# a manual, human-written replacement - layered on top of the restored
# names below.
export_colname_restore <- setNames(df_raw_original_colnames, colnames(df_raw))
export_colname_overrides <- c(
  "Intervention.ID"                  = "Intervention ID",
  "Row_ID"                           = "Row ID",
  "Study.ID"                         = "Study ID",
  "Source.of.article"                = "Source of article",
  "Short.Title"                      = "Short Title",
  "World.Region"                     = "World Region",
  "Sample.Size..analytical."         = "Sample Size (analytical)",
  "Study.Population"                 = "Study Population",
  "Age.Range"                        = "Age Range",
  "Mean.Age"                         = "Mean Age",
  "Unit.for.Age"                     = "Unit for Age",
  "Age.category"                     = "Age category",
  "Intervention.Type"                = "Intervention Type",
  "Specific.intervention.type"       = "Specific intervention type",
  "Intervention.Name"                = "Intervention Name",
  "Intervention.Target"              = "Intervention vict/perp",
  "Level.location.of.intervention"   = "Level/location of intervention",
  "Description.of.the.Intervention"  = "Description of the Intervention",
  "Primary.INSPIRE.Category"         = "Primary INSPIRE Category",
  "Second.INSPIRE.Category"          = "Second INSPIRE Category",
  "Study.Design"                     = "Study Design"
)
export_colname_lookup <- export_colname_restore
export_colname_lookup[names(export_colname_overrides)] <- export_colname_overrides

# The columns of the Excel file. CSV downloads are trimmed to these so the
# working columns the app builds for itself do not end up in the files
# people download. evidence_rating is deliberately left out: it's the same
# study-level effectiveness value as "Effectiveness.v3...Effective." (see
# clean_export_df()), just worded differently - keeping both would show
# dashboard users two columns that say the same thing in different words.
# The rest of this list is internal/admin coding fields or columns judged
# not useful to a dashboard user, dropped at the user's request.
csv_export_columns_dropped <- c(
  "Code", "Lastly.updated",
  "Control.Intervention..for.studies.with.control.groups.",
  "Intervention.length..Numerical.",
  "Intervention.length..unit.",
  "Frequency",
  "Intensity",
  "Who.delivered",
  "Study.Aims",
  "No.of.clusters.randomised.to.intervention..Cluster.RCTs.",
  "No.of.clusters.randomised.to.control..Cluster.RCTs.",
  "Data.Collection.Methods",
  "Reports.ethical.clearance.with.an.IRB.",
  "Mentions.having.a.response.plan.in.place..including.providing.referral.information.or.an.adverse.event.protocol.for.acute.cases.where.participants.require.follow.ups.",
  "Vic.Perp..Intervention.level",
  "Other.VAC.outcomes",
  "Pre.2018.evaluation.vic.perp.outcomes"
)
csv_export_columns <- setdiff(colnames(df_raw), csv_export_columns_dropped)
df <- df_raw

# Display order for INSPIRE strategies
inspire_order <- rev(c(
  "Implementation and enforcement of laws",
  "Norms and values",
  "Safe environments",
  "Parent and caregiver support",
  "Income and economic strengthening",
  "Response and support services",
  "Education and life skills"
))


# =============================================================================
# 4. Data standardisation and cleaning helpers
# =============================================================================

# Tidy a multi-value age-group string into a comma-separated unique list
clean_age_groups <- function(x) {
  x <- as.character(x)
  x <- gsub("[\r\n\t]+", "|", x)
  x <- gsub("^\\|+|\\|+$", "", x)
  parts <- unlist(strsplit(x, "\\|", fixed = FALSE))
  parts <- trimws(parts)
  parts <- parts[nzchar(parts)]
  parts <- unique(parts)
  if (length(parts) == 0) return(NA_character_)
  paste(parts, collapse = ", ")
}

# -----------------------------------------------------------------------------
# Map the (new) extraction schema onto the legacy column names used throughout
# the app: aggregate the per-outcome columns and rename / back-fill the fields
# the rest of the app expects.
# -----------------------------------------------------------------------------
standardise_new_schema <- function(d) {
  
  nms <- names(d)
  
  # --- Identify the repeated per-outcome columns -----------------------------
  outcome_domain_cols     <- grep("^Outcome\\.[0-9]+\\.Domain$", nms, value = TRUE)
  reported_by_cols        <- grep("^Cleaned\\.?Outcome\\.[0-9]+\\.reported\\.by\\.1$", nms, value = TRUE)
  findings_cols           <- grep("^Summary\\.of\\.Key\\.Findings\\.\\.Outcome\\.[0-9]+\\.$", nms, value = TRUE)
  subgroup_conducted_cols <- grep("^Outcome\\.[0-9]+\\.Subgroup\\.Analysis\\.Conducted\\.$", nms, value = TRUE)
  subgroup_groups_cols    <- grep("^Outcome\\.[0-9]+\\.Subgroup\\.Analysis\\.Which\\.groups\\.$", nms, value = TRUE)
  specific_outcome_cols   <- grep("^Cleaned\\.specific\\.Outcome\\.[0-9]+$", nms, value = TRUE)
  
  # Helper: collapse a set of columns row-wise into a unique "; "-joined string
  collapse_unique_row <- function(d, cols) {
    apply(d[, cols, drop = FALSE], 1, function(x) {
      x <- unique(trimws(as.character(x)))
      x <- x[!is.na(x) & nzchar(x)]
      if (length(x) == 0) return(NA_character_)
      paste(x, collapse = "; ")
    })
  }
  
  
  if (length(specific_outcome_cols) > 0)
    d$Specific.Outcome <- collapse_unique_row(d, specific_outcome_cols)
  
  if (length(outcome_domain_cols) > 0)
    d$Outcome.Domain <- collapse_unique_row(d, outcome_domain_cols)
  
  if (length(reported_by_cols) > 0)
    d$Reported.by <- collapse_unique_row(d, reported_by_cols)
  
  if (length(subgroup_conducted_cols) > 0)
    d$Sub.group.analyses.conducted. <- collapse_unique_row(d, subgroup_conducted_cols)
  
  if (length(subgroup_groups_cols) > 0)
    d$Which.subgroups. <- collapse_unique_row(d, subgroup_groups_cols)
  
  # Findings: label each non-empty finding with its outcome number
  if (length(findings_cols) > 0) {
    d$Summary.of.Key.Findings <- apply(d[, findings_cols, drop = FALSE], 1, function(row) {
      parts <- c()
      for (i in seq_along(row)) {
        val <- trimws(as.character(row[i]))
        if (!is.na(val) && nzchar(val)) {
          parts <- c(parts, paste0("Outcome ", i, ": ", val))
        }
      }
      if (length(parts) == 0) return(NA_character_)
      paste(parts, collapse = "\n\n")
    })
  }
  
  # --- Rename / map new-schema fields onto legacy names ----------------------
  if (!"Author" %in% names(d) && "Short.Title" %in% names(d))
    d$Author <- d$Short.Title
  
  if (!"Study.Title" %in% names(d) && "Title" %in% names(d))
    d$Study.Title <- d$Title
  
  if (!"Age.Group.Description" %in% names(d)) {
    d$Age.Group.Description <- if ("Age.category" %in% names(d)) d$Age.category else NA_character_
  }
  d$Age.Group.Description <- sapply(d$Age.Group.Description, clean_age_groups)
  
  if (!"Sample.Size" %in% names(d) && "Sample.Size..analytical." %in% names(d))
    d$Sample.Size <- d$Sample.Size..analytical.
  
  if (!"Age..unit." %in% names(d) && "Unit.for.Age" %in% names(d))
    d$Age..unit. <- d$Unit.for.Age
  
  if (!"Intervention.Length" %in% names(d) && "Intervention.length..Numerical." %in% names(d))
    d$Intervention.Length <- d$Intervention.length..Numerical.
  
  if (!"Intervention.Length..unit." %in% names(d) && "Intervention.length..unit." %in% names(d))
    d$Intervention.Length..unit. <- d$Intervention.length..unit.
  
  if (!"Description.of.Intervention" %in% names(d) && "Description.of.the.Intervention" %in% names(d))
    d$Description.of.Intervention <- d$Description.of.the.Intervention
  
  if (!"Who.delivered.the.intervention." %in% names(d) && "Who.delivered" %in% names(d))
    d$Who.delivered.the.intervention. <- d$Who.delivered
  
  if (!"Setting" %in% names(d) && "Level.location.of.intervention" %in% names(d))
    d$Setting <- d$Level.location.of.intervention
  
  if (!"Income.Classification" %in% names(d)) d$Income.Classification <- NA_character_
  if (!"World.Region" %in% names(d))          d$World.Region <- NA_character_
  if (!"Lastly.updated" %in% names(d))        d$Lastly.updated <- NA_character_
  
  # --- INSPIRE fallback detection --------------------------------------------
  if (!"Primary.INSPIRE.Category" %in% names(d)) {
    cand <- grep("Primary.*INSPIRE", names(d), value = TRUE, ignore.case = TRUE)[1]
    if (!is.na(cand)) d$Primary.INSPIRE.Category <- d[[cand]]
  }
  if (!"Second.INSPIRE.Category" %in% names(d)) {
    cand <- grep("Second.*INSPIRE", names(d), value = TRUE, ignore.case = TRUE)[1]
    if (!is.na(cand)) d$Second.INSPIRE.Category <- d[[cand]]
  }
  
  # --- Final data compatibility mappings -------------------------------------
  # Level.location.of.intervention is the delivery/location/context field.
  if (!"Level.location.of.intervention" %in% names(d))
    d$Level.location.of.intervention <- NA_character_
  
  # Fields used by the Studies included table but not present in final data.
  if (!"Methodology" %in% names(d)) d$Methodology <- NA_character_
  if (!"Measure" %in% names(d))     d$Measure <- NA_character_
  
  if (!"Recipient.of.Intervention" %in% names(d)) {
    d$Recipient.of.Intervention <- if ("Intervention.Target" %in% names(d)) d$Intervention.Target else NA_character_
  }
  
  if (!"Adaptations.for.diverse.groups" %in% names(d))
    d$Adaptations.for.diverse.groups <- NA_character_
  
  if (!"Primary.or.Secondary.Outcome." %in% names(d))
    d$Primary.or.Secondary.Outcome. <- NA_character_
  
  # Keep raw intervention-level effectiveness as character.
  d$Effectiveness..intervention <- if ("Effectiveness..intervention" %in% names(d)) {
    as.character(d$Effectiveness..intervention)
  } else {
    NA_character_
  }
  
  d
}

# Collapse whitespace / line breaks in a free-text field
clean_text_field <- function(x) {
  x <- as.character(x)
  x <- gsub("[\r\n\t]+", " ", x)
  x <- gsub("\\s+", " ", x)
  x <- trimws(x)
  x[x == ""] <- NA_character_
  x
}

# Given several candidate strings for the same field across a group of rows
# (e.g. multiple evaluations of the same intervention), pick the most
# informative one rather than just the first. "More informative" = more
# comma/semicolon/"and"-separated components (e.g. "Male and Female" beats
# "Female"), with string length as a tie-breaker.
pick_richest_value <- function(x) {
  x <- as.character(x)
  x <- trimws(x)
  x <- x[!is.na(x) & nzchar(x) & !tolower(x) %in% c("na", "n/a", "not reported", "none")]
  if (length(x) == 0) return(NA_character_)
  
  n_components <- vapply(x, function(v) {
    length(unlist(strsplit(v, "(?i)\\s*(,|;|/|\\band\\b)\\s*", perl = TRUE)))
  }, integer(1))
  
  x[order(-n_components, -nchar(x))][1]
}

# Standardise outcome-domain labels to a consistent vocabulary and order
standardise_outcome_domain <- function(x) {
  x <- as.character(x)
  x <- gsub("[\r\n\t]+", " ", x)
  x <- gsub("\\s+", " ", x)
  x <- trimws(x)
  
  parts <- strsplit(x, "\\s*[;,|]\\s*", perl = TRUE)
  
  parts <- lapply(parts, function(v) {
    v <- trimws(v)
    v <- v[nzchar(v)]
    
    v <- dplyr::case_when(
      str_detect(tolower(v), "victim")        ~ "Victimisation",
      str_detect(tolower(v), "perpetr")       ~ "Perpetration",
      str_detect(tolower(v), "both|combined") ~ "Both (combined vic/perp outcome)",
      TRUE ~ NA_character_
    )
    v <- v[!is.na(v)]
    
    order_vec <- c("Perpetration", "Victimisation",
                   "Both (combined vic/perp outcome)")
    v <- intersect(order_vec, v)
    
    if (length(v) == 0) return(NA_character_)
    paste(v, collapse = "; ")
  })
  
  unlist(parts)
}

# Split a delimited string into a cleaned vector of unique tokens
clean_tokens <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x <- gsub("[\r\n\t]+", ";", x)
  x <- gsub("\\s*\\|\\s*", ";", x)
  x <- gsub(";+", ";", x)
  
  parts <- strsplit(x, ";", perl = TRUE)
  lapply(parts, function(v) {
    v <- trimws(v)
    v <- v[nzchar(v)]
    unique(v)
  })
}

# Split a multi-value field on a range of delimiters into a list column
parse_multi_values <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x <- gsub("[\r\n\t]+", "|", x)
  x <- gsub("\\s*;\\s*", "|", x)
  x <- gsub("\\s*,\\s*", "|", x)
  x <- gsub("\\|+", "|", x)
  x <- gsub("^\\|+|\\|+$", "", x)
  
  out <- strsplit(x, "\\|", perl = TRUE)
  lapply(out, function(v) {
    v <- trimws(v)
    v <- v[nzchar(v)]
    unique(v)
  })
}


# --- Apply the standardisation pipeline --------------------------------------
df <- standardise_new_schema(df)
df$Outcome.Domain <- standardise_outcome_domain(df$Outcome.Domain)

# Tidy key free-text columns.
# NOTE: "Country" and "World.Region" are deliberately excluded here -
# clean_text_field() collapses newlines to a plain space, which would turn a
# multi-value cell into one unreadable run-on blob before the dedicated
# standardisation blocks below (which handle multi-value cells properly) get
# a chance to run.
cols_to_clean <- c(
  "Intervention.Name",
  "Study.Title",
  "Outcome.Domain",
  "Specific.Outcome",
  "Level.location.of.intervention"
)

for (col in intersect(cols_to_clean, names(df))) {
  df[[col]] <- clean_text_field(df[[col]])
}


# =============================================================================
# 5. Derived list columns and dropdown choices
# =============================================================================

# List columns used for filtering
df$Outcome_Domain_List   <- clean_tokens(df$Outcome.Domain)
df$Specific_Outcome_List <- clean_tokens(df$Specific.Outcome)
df$Reported_By_List      <- clean_tokens(df$Reported.by)

# Dropdown choices derived from the list columns
outcome_domain_choices   <- sort(unique(unlist(df$Outcome_Domain_List)))
specific_outcome_choices <- sort(unique(unlist(df$Specific_Outcome_List)))

# Age groups
df$Age.Group.Description <- as.character(df$Age.Group.Description)
df$Age.Groups_List       <- parse_multi_values(df$Age.Group.Description)

# Canonical age-group order (youngest to oldest, not alphabetical) - matches
# the order already used for the age-group badges shown in the tables.
age_group_order <- c(
  "Early Childhood (0-4 years)",
  "Middle childhood (5-9 years)",
  "Adolescence (10-18 years)",
  "Adults (18+ years)",
  "NA"
)

# Canonical display order for intervention locations (settings): known
# locations first, in this order, then any others alphabetically.
context_order <- c(
  "School", "Community", "Online", "Online at school", "Health", "Home",
  "Residential/alt care", "Workplace", "Religion/faith", "Policy or law",
  "Criminal justice system", "Other"
)

# Classify free-text age-group labels into the four canonical age groups (or
# "NA" when not reported / unrecognised). Vectorised.
classify_age_group <- function(labels) {
  lab <- tolower(trimws(as.character(labels)))
  dplyr::case_when(
    str_detect(lab, "^na$|^n/a$|not reported|not available|missing") ~ "NA",
    str_detect(lab, "early childhood|0\\s*[-\u2013]\\s*4|0\\sto\\s4") ~ "Early Childhood (0-4 years)",
    str_detect(lab, "middle childhood|5\\s*[-\u2013]\\s*9|5\\sto\\s9") ~ "Middle childhood (5-9 years)",
    str_detect(lab, "adolesc|10\\s*[-\u2013]\\s*18|10\\sto\\s18|teen|youth") ~ "Adolescence (10-18 years)",
    str_detect(lab, "adult|18\\+|18 and over|18\\s*[-\u2013]\\s*\\+") ~ "Adults (18+ years)",
    TRUE ~ "NA"
  )
}

# Order a vector of age-group labels: known categories first (in age order),
# then any other/unexpected labels alphabetised at the end.
order_age_groups <- function(labels) {
  labels <- unique(trimws(as.character(labels)))
  labels <- labels[nzchar(labels)]
  c(intersect(age_group_order, labels), sort(setdiff(labels, age_group_order)))
}

# Level / location of intervention -> list column
df$Level.location.of.intervention <- as.character(df$Level.location.of.intervention)
df$Level.location.of.intervention <- gsub("[\r\n\t]+", "|", df$Level.location.of.intervention)
df$Level.location.of.intervention <- trimws(df$Level.location.of.intervention)
df$Level_List <- strsplit(df$Level.location.of.intervention, "\\|")
df$Level_List <- lapply(df$Level_List, function(x) {
  x <- trimws(x)
  x[nzchar(x)]
})

# --- INSPIRE category list column (Prevention explorer) ---------------------
# Combines Primary + Second INSPIRE category per row, normalising known typos
# ("Implemention...") and stray line breaks, and dropping "N/A"-style
# non-values so only genuine categories are ever shown or filterable.
clean_inspire_category <- function(x) {
  raw <- trimws(gsub("[\r\n\t]+", " ", as.character(x)))
  lab <- tolower(raw)
  dplyr::case_when(
    is.na(x) | !nzchar(raw) | lab %in% c("na", "n/a", "not applicable", "none") ~ NA_character_,
    str_detect(lab, "implement")    ~ "Implementation and enforcement of laws",
    str_detect(lab, "norms")        ~ "Norms and values",
    str_detect(lab, "safe environ") ~ "Safe environments",
    str_detect(lab, "parent")       ~ "Parent and caregiver support",
    str_detect(lab, "income")       ~ "Income and economic strengthening",
    str_detect(lab, "response")     ~ "Response and support services",
    str_detect(lab, "education")    ~ "Education and life skills",
    TRUE ~ raw
  )
}

.inspire_primary_clean <- if ("Primary.INSPIRE.Category" %in% names(df)) {
  clean_inspire_category(df$Primary.INSPIRE.Category)
} else {
  rep(NA_character_, nrow(df))
}
.inspire_second_clean <- if ("Second.INSPIRE.Category" %in% names(df)) {
  clean_inspire_category(df$Second.INSPIRE.Category)
} else {
  rep(NA_character_, nrow(df))
}

df$Inspire_List <- Map(function(a, b) {
  v <- c(a, b)
  v <- v[!is.na(v) & nzchar(v)]
  unique(v)
}, .inspire_primary_clean, .inspire_second_clean)

# Canonical INSPIRE order (I-N-S-P-I-R-E) - reuses the chart ordering above,
# which is stored reversed (for coord_flip bar charts).
inspire_category_order <- rev(inspire_order)

inspire_choices <- sort(unique(unlist(df$Inspire_List)))
inspire_choices <- c(
  intersect(inspire_category_order, inspire_choices),
  sort(setdiff(inspire_choices, inspire_category_order))
)

# --- Country name standardisation --------------------------------------------
# Source cells can list several countries separated by newlines (e.g. a single
# multi-site study). Newlines are converted to a real ", " delimiter (not a
# bare space) so individual country names stay distinguishable downstream.
df$Country <- as.character(df$Country)
df$Country <- gsub("[\r\n\t]+", ", ", df$Country)

# Known data-entry quirk: a couple of rows have "Kenya South Africa" typed as
# free text with no delimiter between the two country names at all. Insert
# one so it's treated like every other multi-country cell below.
df$Country <- gsub("(?i)\\bKenya\\s+South Africa\\b", "Kenya, South Africa", df$Country, perl = TRUE)
df$Country <- gsub("(?i)\\bSouth Africa\\s+Kenya\\b", "South Africa, Kenya", df$Country, perl = TRUE)

df$Country <- gsub(",\\s*,", ", ", df$Country)
df$Country <- gsub("^,\\s*|,\\s*$", "", df$Country)
df$Country <- gsub("[ \t]+", " ", df$Country)
df$Country <- trimws(df$Country)
df$Country[df$Country == ""] <- NA_character_

df$Country[df$Country %in% c("USA", "U.S.A.", "United States")] <- "United States of America"
df$Country[df$Country %in% c("United Kingdom", "UK", "U.K.", "Great Britain")] <- "United Kingdom"
df$Country[grepl("^multi\\b", tolower(df$Country))] <- "Multi-country"

# Canonical lower-case country key (used for map joins and filtering)
df$cn <- tolower(trimws(df$Country))

# --- World Region standardisation --------------------------------------------
# A handful of rows list every region (one per line) for genuinely global /
# multi-region studies. Before collapsing those to the "Multi-region" display
# label below, keep the individual regions they actually cover in a list
# column, so e.g. selecting "South Asia" still surfaces a study that spans
# South Asia plus five other regions, instead of only matching studies whose
# *entire* region cell is the single string "South Asia".
split_region_tokens <- function(x) {
  x <- as.character(x)
  if (is.na(x) || !nzchar(trimws(x))) return(character(0))
  tok <- unlist(strsplit(x, "[\r\n]+"))
  tok <- trimws(tok)
  tok[nzchar(tok)]
}

df$World.Region_List <- lapply(as.character(df$World.Region), function(x) {
  toks <- split_region_tokens(x)
  if (length(toks) > 1) c(toks, "Multi-region") else toks
})

# Collapse the display column itself to the "Multi-region" label already used
# elsewhere in the data, instead of leaving a run-on, unreadable blob of every
# region name glued together.
df$World.Region <- as.character(df$World.Region)
df$World.Region <- ifelse(
  grepl("[\r\n]", df$World.Region),
  "Multi-region",
  df$World.Region
)
df$World.Region <- gsub("[\t]+", " ", df$World.Region)
df$World.Region <- gsub("\\s+", " ", df$World.Region)
df$World.Region <- trimws(df$World.Region)
df$World.Region[df$World.Region == ""] <- NA_character_

# Split a (possibly multi-country) Country cell into individual country names.
# Tolerates the ", " delimiter produced above as well as legacy "|"/";" ones.
split_country_tokens <- function(x) {
  x <- as.character(x)
  if (is.na(x) || !nzchar(trimws(x))) return(character(0))
  x <- gsub("[|;]", ",", x)
  tok <- unlist(strsplit(x, ",", fixed = TRUE))
  tok <- trimws(tok)
  tok[nzchar(tok)]
}

# Number of distinct countries in a vector of Country cells. Multi-country cells
# (e.g. "Kenya, South Africa") are split into their countries, and the generic
# "Multi-country" label (a global study with no named countries) is not a
# country, so it is not counted.
count_countries <- function(country_cells) {
  toks <- unique(unlist(lapply(na.omit(country_cells), split_country_tokens)))
  sum(!grepl("^multi", tolower(toks)))
}

# Individual country choices for filter dropdowns (a combined cell like
# "Kenya, South Africa" contributes both "Kenya" and "South Africa", not one
# combined chunk).
country_choices <- sort(unique(unlist(lapply(na.omit(df$Country), split_country_tokens))))

# Row-wise test: does this (possibly multi-country) Country cell contain any
# of the selected countries? Used instead of `Country %in% selected`, which
# only matches a cell if its *entire* combined string equals one selection.
country_matches_selection <- function(country_cell, selected) {
  vapply(country_cell, function(x) length(intersect(split_country_tokens(x), selected)) > 0, logical(1))
}

# Normalise country names specifically for the World Bank join
normalise_country_for_wb <- function(x) {
  x <- as.character(x)
  x <- trimws(x)
  x[x == ""] <- NA_character_
  
  # Detect multi-country cells *before* the name corrections below, since some
  # of those corrections introduce their own comma (e.g. "Congo, Dem. Rep.").
  is_multi <- !is.na(x) & vapply(x, function(v) length(split_country_tokens(v)) > 1, logical(1))
  
  x[x %in% c("UK", "U.K.", "United Kingdom", "Great Britain")] <- "United Kingdom"
  x[x %in% c("USA", "U.S.A.", "United States of America", "United States")] <- "United States"
  x[x %in% c("Democratic Republic of Congo", "DRC", "DR Congo", "Congo DRC")] <- "Congo, Dem. Rep."
  x[x %in% c("Timor Leste")] <- "Timor-Leste"
  
  # Multi-country / combined-country rows are not joined to a single polygon
  x[grepl("^multi\\b", tolower(x))] <- NA_character_
  x[is_multi] <- NA_character_
  
  x
}

# Lower-case World Bank join key for a country name: applies the corrections
# above, then maps remaining variants to the World Bank's own spelling (its file
# uses "Russian Federation", "Côte d’Ivoire" and "Türkiye").
wb_join_key <- function(x) {
  key <- tolower(trimws(normalise_country_for_wb(x)))
  key[key %in% "russia"] <- "russian federation"
  key[key %in% c("cote d’ivoire", "côte d’ivoire", "cote d'ivoire")] <- "côte d’ivoire"
  key[key %in% "turkiye"] <- "türkiye"
  key
}

# --- Filter choice vectors ---------------------------------------------------
context_by_filter <- sort(unique(trimws(unlist(df$Level_List))))
context_by_filter <- context_by_filter[nzchar(context_by_filter)]


# =============================================================================
# 6. World Bank income / region join
# =============================================================================

wb_path <- "World Bank Classification.xlsx"

if (file.exists(wb_path)) {
  
  wb <- readxl::read_excel(wb_path)
  names(wb) <- make.names(names(wb))
  
  # Expected columns after make.names(): Country, Region, Income.group
  if (!"Country" %in% names(wb))      stop("World Bank file is missing a 'Country' column")
  if (!"Income.group" %in% names(wb)) stop("World Bank file is missing an 'Income group' column")
  if (!"Region" %in% names(wb))       stop("World Bank file is missing a 'Region' column")
  
  # Build join keys using normalised country names
  df <- df %>%
    mutate(
      Country    = as.character(Country),
      Country_wb = normalise_country_for_wb(Country),
      df_cn      = wb_join_key(Country)
    )
  
  wb <- wb %>%
    mutate(
      Country = as.character(Country),
      wb_cn   = tolower(trimws(Country))
    )
  
  # Join and back-fill the dashboard columns from the WB file
  df_joined <- df %>%
    left_join(wb %>% select(wb_cn, Region, Income.group),
              by = c("df_cn" = "wb_cn")) %>%
    mutate(
      Income.Classification = ifelse(
        !is.na(Income.group) & nzchar(trimws(Income.group)),
        Income.group, Income.Classification),
      # World.Region already comes populated from the extraction (using its
      # own region taxonomy); only fall back to the World Bank region when
      # it's genuinely missing, rather than overwriting good existing data.
      World.Region = ifelse(
        !is.na(World.Region) & nzchar(trimws(as.character(World.Region))),
        as.character(World.Region),
        ifelse(!is.na(Region) & nzchar(trimws(Region)), Region, World.Region)
      )
    ) %>%
    select(-Income.group, -Region)
  
  df <- df_joined
  
  # Record any countries that failed to match (written for inspection)
  unmatched <- df %>%
    mutate(Country = normalise_country_for_wb(Country)) %>%
    filter(!is.na(Country)) %>%
    filter(is.na(Income.Classification) | !nzchar(trimws(Income.Classification))) %>%
    distinct(Country) %>%
    arrange(Country)
  
  if (nrow(unmatched) > 0) {
    warning("World Bank join: unmatched countries detected: ",
            paste(head(unmatched$Country, 20), collapse = ", "),
            if (nrow(unmatched) > 20)
              paste0(" ... (+", nrow(unmatched) - 20, " more)") else "")
    try(write.csv(unmatched, "unmatched_countries_worldbank.csv",
                  row.names = FALSE), silent = TRUE)
  }
  
} else {
  warning("World Bank Classification.xlsx not found in app directory; ",
          "Income.Classification may be incomplete.")
}


# =============================================================================
# 7. Defensive coercions, evidence-rating helpers and badge helpers
# =============================================================================

# --- Coerce expected character columns ---------------------------------------
safe_char_cols <- c(
  "Reported.by", "Specific.Outcome", "Outcome.Domain",
  "Primary.INSPIRE.Category", "Second.INSPIRE.Category",
  "Age.Group.Description", "Income.Classification", "Study.ID", "Author",
  "Study.Title", "Intervention.Name", "Setting", "Summary.of.Key.Findings",
  "World.Region", "Country", "Study.Design", "Study.Aims", "Year",
  "Lastly.updated", "Sub.group.analyses.conducted.", "Which.subgroups."
)
for (cname in intersect(safe_char_cols, names(df))) {
  df[[cname]] <- as.character(df[[cname]])
}

if ("evidence_rating" %in% colnames(df))   df$evidence_rating   <- as.character(df$evidence_rating)
if ("Country" %in% colnames(df))           df$Country           <- as.character(df$Country)
if ("Intervention.Name" %in% colnames(df)) df$Intervention.Name <- trimws(as.character(df$Intervention.Name))
if ("Study.ID" %in% colnames(df))          df$Study.ID          <- as.character(df$Study.ID)
if ("publication_year" %in% colnames(df))
  df$publication_year <- suppressWarnings(as.integer(as.character(df$publication_year)))

# --- Evidence-rating classifiers ---------------------------------------------

# Effectiveness classifier
get_evidence_rating_v3 <- function(x) {
  z <- tolower(trimws(as.character(x)))
  z <- gsub("\\s+", " ", z)
  dplyr::case_when(
    z %in% c("reduced", "reduceed")                         ~ "Reduced CSV",
    z %in% c("no change", "nochange", "no effect")          ~ "No change in CSV",
    z %in% c("increase", "increased")                       ~ "Increased CSV",
    z %in% c("not categorised", "not categorized", "na", "n/a") ~ "Not categorised",
    TRUE ~ "Not categorised"
  )
}

# Resolve the effect codes recorded for one outcome (in one domain) to a single
# direction, using the same rule as the overall evidence ratings: any increase
# wins (I), otherwise any reduction (R), otherwise no change (NC). Cells such as
# "R + NC" are split, and anything other than R / NC / I is ignored (NA).
resolve_outcome_effect <- function(codes) {
  codes <- trimws(unlist(strsplit(toupper(as.character(codes)), "\\s*\\+\\s*")))
  if ("I" %in% codes) {
    "I"
  } else if ("R" %in% codes) {
    "R"
  } else if ("NC" %in% codes) {
    "NC"
  } else {
    NA_character_
  }
}

# One line per cleaned specific outcome and outcome domain (e.g. "Forced sex
# (Perpetration)") across every row of d - a study's, or a country's-worth of,
# evaluation rows. Repeat measurements of the same outcome and domain, whether
# in other outcome slots or other rows, are collapsed to a single effect code
# with resolve_outcome_effect(). Returns a data frame with the columns label and
# code (NA when no effect was recorded), in order of first appearance.
outcome_effect_lines <- function(d) {
  outcome_cols <- paste0("Cleaned.specific.Outcome.", 1:9)
  domain_cols  <- paste0("Outcome.", 1:9, ".Domain")
  effect_cols  <- paste0("Outcome..", 1:9, "..Outcome.effect.category")

  labels <- character(0)
  codes  <- list()

  for (i in seq_len(nrow(d))) {
    row_df <- d[i, , drop = FALSE]
    for (k in seq_along(outcome_cols)) {
      oc <- outcome_cols[k]
      if (!oc %in% names(row_df)) next

      raw_names <- as.character(row_df[[oc]])
      raw_names <- gsub("[\r\n\t]+", " ", raw_names)
      names_k <- trimws(unlist(strsplit(raw_names, "\\s*[;|]\\s*", perl = TRUE)))
      names_k <- names_k[!is.na(names_k) & nzchar(names_k)]
      names_k <- names_k[!tolower(names_k) %in% c("na", "n/a", "not reported", "missing")]
      if (length(names_k) == 0) next

      dc <- domain_cols[k]
      domain <- if (dc %in% names(row_df)) {
        standardise_outcome_domain(row_df[[dc]])
      } else {
        NA_character_
      }
      ec <- effect_cols[k]
      code_raw <- if (ec %in% names(row_df)) as.character(row_df[[ec]]) else NA_character_

      for (nm in names_k) {
        label <- if (is.na(domain)) nm else paste0(nm, " (", domain, ")")
        if (!label %in% labels) labels <- c(labels, label)
        codes[[label]] <- c(codes[[label]], code_raw)
      }
    }
  }

  data.frame(
    label = labels,
    code  = vapply(labels, function(l) resolve_outcome_effect(codes[[l]]),
                   FUN.VALUE = character(1), USE.NAMES = FALSE),
    stringsAsFactors = FALSE
  )
}

# Standardise income classification to HIC / MIC / LIC / Multicountry
standardise_income <- function(x) {
  x2 <- tolower(as.character(x))
  dplyr::case_when(
    str_detect(x2, "multi") ~ "Multicountry",
    str_detect(x2, "\\bhic\\b") |
      str_detect(x2, "high income") |
      str_detect(x2, "high-income") ~ "High Income",
    str_detect(x2, "\\blic\\b") |
      str_detect(x2, "low-income") |
      str_detect(x2, "low income") |
      str_detect(x2, "^low$") ~ "Low Income",
    str_detect(x2, "\\bumic\\b") |
      str_detect(x2, "upper middle") |
      str_detect(x2, "upper-middle") |
      str_detect(x2, "\\blmic\\b") |
      str_detect(x2, "lower middle") |
      str_detect(x2, "lower-middle") |
      str_detect(x2, "middle income") ~ "Middle Income",
    TRUE ~ NA_character_
  )
}

# --- Badge HTML helpers ------------------------------------------------------

# Text colour for a badge, following the Brand Style Guide's accessible colour
# combinations (p. 22): white text on Green, Blue, Purple and Deep Ocean;
# Charcoal text on Orange, Pink and Sand (white on Orange/Pink is too low
# contrast for small text).
badge_text_colour <- function(fill) {
  if (toupper(fill) %in% toupper(c(sfh_orange, sfh_pink, sfh_sand))) sfh_charcoal else "#FFFFFF"
}

# Fixed value -> brand colour lookup so a given category always renders in
# the same colour, no matter which/how many other values sit alongside it in
# a given row (previously colour was assigned by the value's position within
# that row's own list, so the same value could get a different colour - or
# always the first palette colour - depending on what else was present).
badge_colour_for <- function(label, palette_key, badge_palette) {
  pos <- match(label, palette_key)
  if (is.na(pos)) {
    # Deterministic fallback for a value outside the canonical list, so a
    # repeat render of the same "unlisted" value still lands on the same
    # colour instead of depending on row position.
    pos <- (sum(utf8ToInt(label)) %% length(badge_palette)) + 1
  }
  badge_palette[(pos - 1) %% length(badge_palette) + 1]
}

# Render a vector of labels as a wrapped row of coloured pill badges
make_tag_badges_html <- function(x, order = NULL, special_fill = NULL,
                                 special_text = NULL, max_height = NULL) {
  x <- as.character(x)
  x <- x[!is.na(x)]
  x <- stringr::str_squish(x)
  x <- x[nzchar(x)]
  x <- unique(x)
  
  if (!is.null(order)) x <- c(intersect(order, x), setdiff(x, order))
  if (length(x) == 0) return(NA_character_)
  
  badge_palette <- c(sfh_green, sfh_blue, sfh_purple, sfh_orange, sfh_ocean, sfh_pink)
  
  # Colour key is the canonical `order` list (or, failing that, the
  # alphabetised set of values actually seen) so a value's colour never
  # depends on which row it happens to appear in.
  palette_key <- if (!is.null(order)) order else sort(unique(x))
  
  badges <- vapply(seq_along(x), function(i) {
    label <- x[i]
    
    fill <- if (!is.null(special_fill) && label %in% names(special_fill)) {
      special_fill[[label]]
    } else {
      badge_colour_for(label, palette_key, badge_palette)
    }
    
    textc <- if (!is.null(special_text) && label %in% names(special_text)) {
      special_text[[label]]
    } else {
      badge_text_colour(fill)
    }
    
    paste0(
      "<span style='display:inline-block; margin:0; ",
      "padding:3px 8px; border-radius:999px; ",
      "background:", fill, "; border:1px solid ", fill, "; ",
      "color:", textc, "; font-family:\"General Sans\"; ",
      "font-size:11px;'>",
      htmltools::htmlEscape(label),
      "</span>"
    )
  }, FUN.VALUE = character(1))
  
  inner <- paste0(
    "<div style='display:flex; flex-wrap:wrap; gap:4px;'>",
    paste(badges, collapse = ""),
    "</div>"
  )
  
  if (!is.null(max_height)) {
    paste0(
      "<div style='max-height:", max_height,
      "px; overflow-y:auto; padding-right:4px;'>",
      inner, "</div>"
    )
  } else {
    inner
  }
}

# Plain-text equivalent of make_tag_badges_html() - same cleaning/ordering,
# joined as readable comma-separated text instead of HTML pill badges. Used
# so CSV exports show "Victimisation, Perpetration" rather than raw <span> markup.
make_tag_plain <- function(x, order = NULL) {
  x <- as.character(x)
  x <- x[!is.na(x)]
  x <- stringr::str_squish(x)
  x <- x[nzchar(x)]
  x <- unique(x)

  if (!is.null(order)) x <- c(intersect(order, x), setdiff(x, order))
  if (length(x) == 0) return(NA_character_)

  paste(x, collapse = ", ")
}

# Render a single pill badge
make_uniform_badge <- function(label, fill, text = badge_text_colour(fill)) {
  paste0(
    "<span style='",
    "display:inline-block;",
    "padding:4px 10px;",
    "border-radius:999px;",
    "background:", fill, ";",
    "color:", text, ";",
    "font-family:\"General Sans\";",
    "font-size:11px;",
    "line-height:1.2;",
    "white-space:nowrap;",
    "'>",
    htmltools::htmlEscape(label),
    "</span>"
  )
}


# =============================================================================
# 8. Derived analytical columns and event logging
# =============================================================================

# --- Publication year and last-updated --------------------------------------
df$publication_year <- suppressWarnings(
  as.integer(round(as.numeric(df$Year), 0))
)

# Turn the Lastly.updated column into R-compatible dates.
parse_lastly_updated_date <- function(x) {
  
  x <- trimws(as.character(x))
  
  # Treat blank values and N/A as missing.
  x[
    is.na(x) |
      tolower(x) %in% c("", "na", "n/a", "none")
  ] <- NA_character_
  
  # Try common written date formats, including:
  # 2026.07.31
  # 2026-07-31
  # 31/07/2026
  parsed <- suppressWarnings(
    as.Date(
      lubridate::parse_date_time(
        x,
        orders = c(
          "ymd",
          "ymd HMS",
          "dmy",
          "dmy HMS",
          "mdy",
          "mdy HMS"
        )
      )
    )
  )
  
  # Also handle Excel serial dates if they appear as five-digit numbers.
  excel_number <- !is.na(x) &
    grepl("^\\d{5}(\\.\\d+)?$", x)
  
  if (any(excel_number)) {
    parsed[excel_number] <- as.Date(
      as.numeric(x[excel_number]),
      origin = "1899-12-30"
    )
  }
  
  parsed
}

# Keep the original text version.
df$lastly_updated <- as.character(df$Lastly.updated)

# Create a proper Date version.
df$lastly_updated_date <- parse_lastly_updated_date(
  df$Lastly.updated
)

# The date used to decide whether something is New.
new_badge_cutoff <- as.Date("2027-01-01")

# For our dashboard, TRUE means the date is later than 1 January 2027.
df$is_new_update <- !is.na(df$lastly_updated_date) &
  df$lastly_updated_date > new_badge_cutoff

# --- Study-level effectiveness (Studies included) ---------------------------
if ("Effectiveness.v3...Effective." %in% names(df)) {
  df$evidence_direction_raw <- as.character(df$Effectiveness.v3...Effective.)
  df$evidence_direction     <- as.character(df$Effectiveness.v3...Effective.)
  df$evidence_rating        <- get_evidence_rating_v3(df$Effectiveness.v3...Effective.)
} else {
  df$evidence_direction_raw <- NA_character_
  df$evidence_direction     <- NA_character_
  df$evidence_rating        <- "Not categorised"
}

# --- Intervention-level effectiveness (Prevention explorer) ----------------
df$intervention_effectiveness_code <- if ("Effectiveness..intervention" %in% names(df)) {
  toupper(trimws(as.character(df$Effectiveness..intervention)))
} else {
  NA_character_
}

df$intervention_effectiveness_label <- dplyr::case_when(
  df$intervention_effectiveness_code == "R"  ~ "Reduced CSV",
  df$intervention_effectiveness_code == "NC" ~ "No change in CSV",
  df$intervention_effectiveness_code == "I"  ~ "Increased CSV",
  df$intervention_effectiveness_code == "NA" ~ "Not categorised",
  TRUE ~ "Not categorised"
)

# --- Standardised income bucket ----------------------------------------------
# Force a "Multicountry" bucket whenever the Country field lists more than one
# country (or is explicitly labelled "Multi..."), even if the World Bank join
# failed / was skipped for that row.
df$income_std <- dplyr::case_when(
  vapply(df$Country, function(v) length(split_country_tokens(v)) > 1, logical(1)) ~ "Multicountry",
  str_detect(tolower(as.character(df$Country)), "\\bmulti") ~ "Multicountry",
  TRUE ~ standardise_income(df$Income.Classification)
)

# --- Year range for sliders --------------------------------------------------
year_vec <- df$publication_year
year_vec <- year_vec[is.finite(year_vec)]

if (length(year_vec) == 0) {
  min_year <- 2000L
  max_year <- as.integer(format(Sys.Date(), "%Y"))
} else {
  min_year <- min(year_vec, na.rm = TRUE)
  max_year <- max(year_vec, na.rm = TRUE)
}
min_year <- as.integer(min_year)
max_year <- as.integer(max_year)

if (!is.finite(min_year) || !is.finite(max_year)) {
  min_year <- 2000L
  max_year <- as.integer(format(Sys.Date(), "%Y"))
}
if (min_year >= max_year) max_year <- min_year + 1L

# --- Fixed income level order used in plots ----------------------------------
income_levels <- c("High Income", "Middle Income", "Low Income", "Multicountry")

# --- Income group of each individual country ---------------------------------
# Used by the income charts. Multi-country cells (e.g. "Kenya, South Africa")
# are split into their individual countries so that every named country is
# counted once, in its own World Bank income group. Only genuinely global
# studies (Country = "Multi-country") stay in the "Multicountry" bucket.
country_income_table <- function(d) {
  d <- d[!is.na(d$Country) & nzchar(trimws(d$Country)), , drop = FALSE]
  tok <- tibble::tibble(cell = unique(as.character(d$Country))) %>%
    dplyr::mutate(country = lapply(cell, split_country_tokens)) %>%
    tidyr::unnest(country)
  
  tok$income_std <- if (exists("wb")) {
    standardise_income(wb$Income.group)[match(wb_join_key(tok$country), wb$wb_cn)]
  } else {
    NA_character_
  }
  tok$income_std[grepl("^multi", tolower(tok$country))] <- "Multicountry"
  
  # Fallback if the World Bank file was unavailable: use the row's own income
  if (anyNA(tok$income_std) && "income_std" %in% names(d)) {
    tok$income_std <- dplyr::coalesce(tok$income_std,
                                      d$income_std[match(tok$country, d$Country)])
  }
  dplyr::distinct(tok, country, income_std)
}

# --- Ensure INSPIRE columns exist --------------------------------------------
inspire_cols <- c("Primary.INSPIRE.Category",
                  "Second.INSPIRE.Category")
for (col in inspire_cols) {
  if (!(col %in% colnames(df))) df[[col]] <- NA_character_
}

# --- Lightweight, anonymous usage logging ------------------------------------
# One line per event (time, event, detail) is appended to metrics_log.csv.
# Nothing personal is recorded: no IP address, no cookies, and no free text -
# only page views, session length, which study / country was opened, which
# (fixed-choice) filters were used, downloads and app errors.
metrics_file <- "metrics_log.csv"

log_event <- function(event_type, detail = NA_character_) {
  # Logging must never break the app, so any failure (e.g. a read-only disk)
  # is ignored.
  try({
    detail <- substr(gsub("[\r\n]+", " ", as.character(detail)), 1, 500)
    row <- data.frame(
      time   = format(Sys.time(), tz = "UTC", usetz = TRUE),
      event  = as.character(event_type),
      detail = detail,
      stringsAsFactors = FALSE
    )
    if (!file.exists(metrics_file)) {
      write.table(row, file = metrics_file, sep = ",",
                  row.names = FALSE, col.names = TRUE, append = FALSE)
    } else {
      write.table(row, file = metrics_file, sep = ",",
                  row.names = FALSE, col.names = FALSE, append = TRUE)
    }
  }, silent = TRUE)
  invisible(TRUE)
}

# Log every unhandled app error (message only, truncated) for debugging.
options(shiny.error = function() {
  log_event("app_error", geterrmessage())
})

# --- Client-side JavaScript --------------------------------------------------
# Handles the table "View" buttons and copy-to-clipboard for the citation box.
#
# How the "View" buttons work (R -> browser -> R):
#   1. R builds each button as plain HTML inside the table cell, with a class
#      (view-btn) so JavaScript can find it and a data-id attribute holding the
#      study or intervention ID:
#        <button class="view-btn" data-id="123">View</button>
#      A button inside a table cell cannot be a Shiny actionButton(), which is
#      why this small script is needed.
#   2. When a session starts, the server sends the message "init-dt-view-button"
#      (session$sendCustomMessage). The handler below runs once and installs a
#      single click listener for the whole page.
#   3. The listener is attached to the page (`document`), not to the buttons, so
#      it keeps working when the table redraws its rows (sorting, searching,
#      changing page).
#   4. On a click it reads the button's data-id and sends it back to R as
#      input$show_study_id. `priority: 'event'` makes Shiny treat every click as
#      new, even when the same study is opened twice in a row.
#   5. On the R side, observeEvent(input$show_study_id, ...) builds the pop-up.
js <- "
// Runs once per session, when the server sends 'init-dt-view-button'.
Shiny.addCustomMessageHandler('init-dt-view-button', function(msg) {
  // One listener for the whole page: reacts to clicks on any current or
  // future element with the class 'view-btn'.
  $(document).on('click', '.view-btn', function() {
    // The study / intervention ID stored in the button's data-id attribute.
    var id = $(this).data('id');
    // Send it to R as input$show_study_id.
    Shiny.setInputValue('show_study_id', id, {priority: 'event'});
  });
});

Shiny.addCustomMessageHandler('copy-citation-to-clipboard', function(msg) {
  var text = msg.text || '';

  function fallbackCopy(text) {
    var textarea = document.createElement('textarea');
    textarea.value = text;
    textarea.setAttribute('readonly', '');
    textarea.style.position = 'fixed';
    textarea.style.top = '-9999px';
    textarea.style.left = '-9999px';
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();

    try {
      document.execCommand('copy');
      Shiny.setInputValue('citation_copy_status', 'copied_fallback', {priority: 'event'});
    } catch (err) {
      Shiny.setInputValue('citation_copy_status', 'failed', {priority: 'event'});
    }

    document.body.removeChild(textarea);
  }

  if (navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(text).then(function() {
      Shiny.setInputValue('citation_copy_status', 'copied', {priority: 'event'});
    }).catch(function() {
      fallbackCopy(text);
    });
  } else {
    fallbackCopy(text);
  }
});
"



# Short Titles shared by more than one Study.ID - the per-study download
# filename appends the full Title for these so two different studies never
# suggest the same file name.
duplicated_short_titles <- local({
  st <- trimws(as.character(df$Short.Title))
  keep <- !is.na(st) & nzchar(st) & !is.na(df$Study.ID)
  pairs <- unique(data.frame(Study.ID = as.character(df$Study.ID)[keep], Short.Title = st[keep]))
  counts <- table(pairs$Short.Title)
  names(counts)[counts > 1]
})


# =============================================================================
# 9. Country-map helper functions
# =============================================================================

# Flatten an sf object / tibble list-columns into a plain character data.frame
normalise_df <- function(d) {
  if (inherits(d, "sf")) d <- sf::st_set_geometry(d, NULL)
  d <- as.data.frame(d, stringsAsFactors = FALSE, check.names = FALSE)
  
  for (nm in names(d)) {
    col <- d[[nm]]
    
    if (is.list(col)) {
      d[[nm]] <- vapply(col, function(x) {
        if (is.null(x)) return(NA_character_)
        if (length(x) == 0) return(NA_character_)
        if (is.atomic(x)) {
          paste(as.character(x), collapse = "; ")
        } else {
          tryCatch(paste(unlist(x), collapse = "; "),
                   error = function(e) NA_character_)
        }
      }, FUN.VALUE = character(1), USE.NAMES = FALSE)
      
    } else if (is.factor(col)) {
      d[[nm]] <- as.character(col)
      
    } else if (inherits(col, "POSIXt") || inherits(col, "Date")) {
      d[[nm]] <- as.character(col)
      
    } else {
      if (is.character(col)) d[[nm]] <- as.character(col)
    }
  }
  
  d
}

# Human-readable label for a raw effect value. The per-outcome effect-category
# columns use letter codes (R / NC / I, or a "+"-joined combination such as
# "R + NC"); the study- and intervention-level effectiveness columns use full
# words instead ("Reduced", "No change", "Increase" / "Increased") - this
# covers both vocabularies case-insensitively. Falls back to the original
# (non-uppercased) value for anything outside them.
export_effect_label <- function(code) {
  raw <- trimws(as.character(code))
  z   <- toupper(raw)
  dplyr::case_when(
    is.na(raw) | raw == ""                          ~ NA_character_,
    z %in% c("R", "REDUCED", "REDUCEED")             ~ "Reduced",
    z %in% c("NC", "NO CHANGE", "NOCHANGE", "NO EFFECT") ~ "No change",
    z %in% c("I", "INCREASE", "INCREASED")           ~ "Increased",
    z == "R + NC"                                    ~ "Reduced + No change",
    z == "R + I"                                     ~ "Reduced + Increased",
    z == "NC + I"                                    ~ "No change + Increased",
    TRUE ~ raw
  )
}

# Shared cleaning for every "download the raw data" handler (the home-page,
# per-country, per-study and Studies-included exports). Called on a data
# frame already subset to intersect(csv_export_columns, names(dd)):
#   - strips leading/trailing whitespace (incl. stray leading newlines) so
#     cells don't render as blank
#   - collapses the 9 repeated per-outcome column-groups (81 raw columns)
#     into 4 reader-facing columns
#   - renames the two effectiveness columns so their scope (study- vs
#     intervention-level) is unambiguous, and decodes their raw R/NC/I codes
clean_export_df <- function(dd, n_outcomes = 9) {
  dd <- normalise_df(dd)

  dd <- dplyr::mutate(dd, dplyr::across(dplyr::where(is.character), function(x) {
    x <- trimws(x)
    x[x == ""] <- NA_character_
    x
  }))

  outcome_raw_cols <- unlist(lapply(seq_len(n_outcomes), function(k) c(
    paste0("Outcome.", k, ".Domain"),
    paste0("Specific.Outcome.", k),
    paste0("Cleaned.specific.Outcome.", k),
    paste0("Outcome.", k, ".Category"),
    paste0("Cleaned.Outcome.", k, ".reported.by.1"),
    paste0("Summary.of.Key.Findings..Outcome.", k, "."),
    paste0("Outcome..", k, "..Outcome.effect.category"),
    paste0("Outcome.", k, ".Subgroup.Analysis.Conducted."),
    paste0("Outcome.", k, ".Subgroup.Analysis.Which.groups.")
  )))
  outcome_raw_cols <- intersect(outcome_raw_cols, names(dd))

  if (length(outcome_raw_cols) > 0) {
    n <- nrow(dd)
    domain_out   <- character(n)
    effect_out   <- character(n)
    subgroup_out <- character(n)

    for (i in seq_len(n)) {
      domain_parts   <- character(0)
      effect_parts   <- character(0)
      subgroup_parts <- character(0)

      for (k in seq_len(n_outcomes)) {
        domain_col   <- paste0("Outcome.", k, ".Domain")
        cleaned_col  <- paste0("Cleaned.specific.Outcome.", k)
        effect_col   <- paste0("Outcome..", k, "..Outcome.effect.category")
        sg_c_col     <- paste0("Outcome.", k, ".Subgroup.Analysis.Conducted.")
        sg_w_col     <- paste0("Outcome.", k, ".Subgroup.Analysis.Which.groups.")

        domain   <- if (domain_col %in% names(dd))   dd[[domain_col]][i]   else NA_character_
        cleaned  <- if (cleaned_col %in% names(dd))   dd[[cleaned_col]][i]  else NA_character_
        effect   <- if (effect_col %in% names(dd))    dd[[effect_col]][i]   else NA_character_
        sg_c     <- if (sg_c_col %in% names(dd))      dd[[sg_c_col]][i]     else NA_character_
        sg_w     <- if (sg_w_col %in% names(dd))      dd[[sg_w_col]][i]     else NA_character_

        has_outcome <- (!is.na(domain) && nzchar(domain)) || (!is.na(cleaned) && nzchar(cleaned))
        if (!has_outcome) next

        label <- paste0("Outcome ", k, ": ")
        domain_txt <- paste(stats::na.omit(c(domain, cleaned)), collapse = " - ")
        domain_parts <- c(domain_parts, paste0(label, domain_txt))

        if (!is.na(effect) && nzchar(effect)) {
          effect_parts <- c(effect_parts, paste0(label, export_effect_label(effect)))
        }

        if (!is.na(sg_c) && tolower(sg_c) == "yes") {
          subgroup_parts <- c(subgroup_parts, paste0(label, "Yes - ", if (!is.na(sg_w)) sg_w else "(group not specified)"))
        } else if (!is.na(sg_c)) {
          subgroup_parts <- c(subgroup_parts, paste0(label, sg_c))
        }
      }

      domain_out[i]   <- if (length(domain_parts))   paste(domain_parts, collapse = " | ")   else NA_character_
      effect_out[i]   <- if (length(effect_parts))   paste(effect_parts, collapse = " | ")   else NA_character_
      subgroup_out[i] <- if (length(subgroup_parts)) paste(subgroup_parts, collapse = " | ") else NA_character_
    }

    dd <- dd[, setdiff(names(dd), outcome_raw_cols), drop = FALSE]
    dd$`Outcome domain (cleaned)`    <- domain_out
    dd$`Outcome effectiveness`       <- effect_out
    dd$`Subgroup analysis conducted` <- subgroup_out
  }

  if ("Effectiveness.v3...Effective." %in% names(dd)) {
    dd <- dplyr::rename(dd, `Study-level effectiveness (overall)` = Effectiveness.v3...Effective.)
    dd$`Study-level effectiveness (overall)` <- export_effect_label(dd$`Study-level effectiveness (overall)`)
  }
  if ("Effectiveness..intervention" %in% names(dd)) {
    dd <- dplyr::rename(dd, `Intervention-level effectiveness (overall; post 2018)` = Effectiveness..intervention)
    dd$`Intervention-level effectiveness (overall; post 2018)` <- export_effect_label(dd$`Intervention-level effectiveness (overall; post 2018)`)
  }

  # Finally, swap every remaining dot/underscore-separated working column
  # name for its normal-looking display name (see export_colname_lookup,
  # built from the Excel file's original header text).
  to_rename <- intersect(names(dd), names(export_colname_lookup))
  names(dd)[match(to_rename, names(dd))] <- unname(export_colname_lookup[to_rename])

  # Column order: intervention name right after Year, and the three
  # outcome columns just before the two overall-effectiveness columns.
  nms <- names(dd)
  if (all(c("Intervention Name", "Year") %in% nms)) {
    nms <- setdiff(nms, "Intervention Name")
    nms <- append(nms, "Intervention Name", after = match("Year", nms))
  }
  outcome_trio <- intersect(
    c("Outcome domain (cleaned)", "Outcome effectiveness", "Subgroup analysis conducted"),
    nms
  )
  anchor <- "Study-level effectiveness (overall)"
  if (length(outcome_trio) > 0 && anchor %in% nms) {
    nms <- setdiff(nms, outcome_trio)
    nms <- append(nms, outcome_trio, after = match(anchor, nms) - 1)
  }
  dd <- dd[, nms, drop = FALSE]

  dd
}

# Write a cleaned export data frame to an .xlsx file. Every character column
# is explicitly given Excel's "Text" cell format, so a value that merely
# looks like a date (e.g. an age range of "10-17") is never silently
# reinterpreted and displayed as a date - numeric columns are left alone so
# they stay sortable/summable as numbers.
write_clean_xlsx <- function(dd, file, sheet = "data") {
  wb <- createWorkbook()
  addWorksheet(wb, sheet)
  writeData(wb, sheet, dd)
  text_cols <- which(vapply(dd, is.character, logical(1)))
  if (length(text_cols) > 0 && nrow(dd) > 0) {
    addStyle(wb, sheet, style = createStyle(numFmt = "TEXT"),
             rows = seq_len(nrow(dd)) + 1, cols = text_cols,
             gridExpand = TRUE, stack = TRUE)
  }
  saveWorkbook(wb, file, overwrite = TRUE)
}

# Strip a string down to a filesystem-safe download-filename fragment:
# letters/digits/spaces only, spaces collapsed to underscores.
clean_filename_part <- function(x) {
  x <- gsub("[^[:alnum:] ]", "", x)
  gsub("\\s+", "_", trimws(x))
}

# Count distinct non-empty values
n_dist_nonempty <- function(x) {
  x <- as.character(x)
  x <- x[!is.na(x)]
  x <- trimws(x)
  x <- x[nzchar(x)]
  length(unique(x))
}

# --- Evidence-rating membership tests (used for map counts) ------------------
is_reduced_csv <- function(x) {
  x <- tolower(trimws(as.character(x)))
  x <- gsub("\\s+", " ", x)
  x %in% c("reduced csv", "effective", "reduced")
}

is_no_effect_csv <- function(x) {
  x <- tolower(trimws(as.character(x)))
  x <- gsub("\\s+", " ", x)
  x %in% c("no change in csv", "no effect on csv", "no effect", "no change", "nochange", "noeffect")
}

is_increased_csv <- function(x) {
  x <- tolower(trimws(as.character(x)))
  x <- gsub("\\s+", " ", x)
  x %in% c("increased csv", "increase", "increased")
}

# Canonical country key used to join study rows to map geometry
country_key <- function(x) {
  x <- as.character(x)
  x <- gsub("[\r\n\t]+", " ", x)
  x <- trimws(x)
  x <- gsub("\\s+", " ", x)
  x_low <- tolower(x)
  x_low <- gsub("\u2019", "'", x_low)
  
  # Multi-country rows are not mapped to a single polygon
  x_low[grepl("^multi\\b", x_low)] <- NA_character_
  
  x_low[x_low %in% c("usa", "u.s.a.", "united states",
                     "united states of america")] <- "united states of america"
  x_low[x_low %in% c("uk", "u.k.", "united kingdom",
                     "great britain")] <- "united kingdom"
  x_low[x_low %in% c("drc", "d.r.c.", "dr congo",
                     "democratic republic of congo",
                     "democratic republic of the congo",
                     "congo drc", "congo, dem. rep.",
                     "dem. rep. congo")] <- "democratic republic of the congo"
  x_low[x_low %in% c("tanzania", "united republic of tanzania",
                     "tanzania, united republic of")] <- "tanzania"
  x_low[x_low %in% c("viet nam")] <- "vietnam"
  x_low[x_low %in% c("cote d'ivoire", "c\u00f4te d'ivoire")] <- "cote d'ivoire"
  x_low[x_low %in% c("turkiye", "t\u00fcrkiye", "turkey")] <- "turkey"
  x_low[x_low %in% c("russian federation")] <- "russia"
  x_low[x_low %in% c("korea, republic of", "republic of korea",
                     "south korea")] <- "south korea"
  x_low[x_low %in% c("iran, islamic republic of", "iran")] <- "iran"
  
  x_low
}

# Split a Country cell into the individual countries to map (e.g. "Kenya, South
# Africa" gives both). The generic "Multi-country" label has no polygon, so it
# gives NA.
split_country_for_map <- function(x) {
  x <- as.character(x)
  if (is.na(x) || !nzchar(trimws(x))) return(NA_character_)
  if (grepl("^multi\\b", tolower(x))) return(NA_character_)
  
  tok <- split_country_tokens(x)
  if (length(tok) == 0) return(NA_character_)
  tok
}

# Build a per-country count summary from a study-level data.frame
country_summary_for_map <- function(d) {
  d <- normalise_df(d)
  
  if (!"Country" %in% names(d))           d$Country <- NA_character_
  if (!"Study.ID" %in% names(d))          d$Study.ID <- NA_character_
  if (!"Intervention.Name" %in% names(d)) d$Intervention.Name <- NA_character_
  if (!"evidence_rating" %in% names(d))   d$evidence_rating <- NA_character_
  
  d %>%
    mutate(
      Country           = as.character(Country),
      Study.ID          = as.character(Study.ID),
      Intervention.Name = as.character(Intervention.Name),
      evidence_rating   = as.character(evidence_rating)
    ) %>%
    rowwise() %>%
    mutate(Country_map = list(split_country_for_map(Country))) %>%
    ungroup() %>%
    tidyr::unnest_longer(Country_map, values_to = "Country_map") %>%
    filter(!is.na(Country_map), nzchar(trimws(Country_map))) %>%
    mutate(.cn = country_key(Country_map)) %>%
    filter(!is.na(.cn) & nzchar(.cn)) %>%
    group_by(.cn) %>%
    summarise(
      Country       = first(na.omit(Country_map)),
      studies       = n_dist_nonempty(Study.ID),
      interventions = n_dist_nonempty(Intervention.Name),
      effective     = n_dist_nonempty(Study.ID[is_reduced_csv(evidence_rating)]),
      .groups = "drop"
    ) %>%
    mutate(
      across(c(studies, interventions, effective),
             ~ replace_na(as.integer(.), 0L))
    )
}

# Obtain a cleaned world geometry from tmap::World.
# Cached after the first call - this involves st_make_valid()/st_transform()
# over the full world polygon set, which is expensive to redo on every map
# render.
.world_sf_cache <- NULL

get_world_sf <- function() {
  if (!is.null(.world_sf_cache)) return(.world_sf_cache)
  
  if (!requireNamespace("tmap", quietly = TRUE)) return(NULL)
  
  w <- tryCatch(tmap::World, error = function(e) NULL)
  if (is.null(w)) return(NULL)
  
  if (!inherits(w, "sf")) {
    w <- tryCatch(sf::st_as_sf(w), error = function(e) NULL)
  }
  if (is.null(w)) return(NULL)
  
  namecol <- intersect(c("name", "NAME", "NAME_LONG", "country", "Country"), names(w))[1]
  if (is.na(namecol)) return(NULL)
  
  w <- w %>%
    dplyr::transmute(
      map_country = as.character(.data[[namecol]]),
      .cn         = country_key(map_country),
      geometry    = geometry
    ) %>%
    dplyr::filter(!is.na(.cn) & nzchar(.cn))
  
  w <- tryCatch(sf::st_make_valid(w), error = function(e) w)
  
  keep <- tryCatch(!sf::st_is_empty(w), error = function(e) rep(TRUE, nrow(w)))
  w <- w[keep, , drop = FALSE]
  
  w <- tryCatch(sf::st_transform(w, 4326), error = function(e) w)
  
  .world_sf_cache <<- w
  w
}

# Render a leaflet choropleth of evaluations by country
render_country_map <- function(summary_df,
                               legend_title = "Number of evaluations",
                               zoom = 2.5) {
  world_sf <- get_world_sf()
  
  # Fallback empty map
  if (is.null(world_sf) || nrow(world_sf) == 0 ||
      is.null(summary_df) || nrow(summary_df) == 0) {
    return(
      leaflet(options = leafletOptions(worldCopyJump = TRUE)) %>%
        addProviderTiles(providers$Esri.WorldGrayCanvas) %>%
        setView(lng = 20, lat = 20, zoom = 2)
    )
  }
  
  summary_df <- as.data.frame(summary_df, stringsAsFactors = FALSE, check.names = FALSE)
  
  if (!".cn" %in% names(summary_df)) {
    summary_df$.cn <- country_key(summary_df$Country)
  } else {
    summary_df$.cn <- country_key(summary_df$.cn)
  }
  
  needed <- c("Country", "studies", "interventions", "effective")
  for (nm in needed) {
    if (!nm %in% names(summary_df)) summary_df[[nm]] <- NA
  }
  
  num_cols <- c("studies", "interventions", "effective")
  for (nm in num_cols) {
    summary_df[[nm]] <- suppressWarnings(as.integer(as.numeric(summary_df[[nm]])))
    summary_df[[nm]][is.na(summary_df[[nm]])] <- 0L
  }
  summary_df$Country <- as.character(summary_df$Country)
  
  summary_df <- summary_df %>%
    dplyr::filter(!is.na(.cn) & nzchar(.cn)) %>%
    dplyr::group_by(.cn) %>%
    dplyr::summarise(
      Country       = dplyr::first(na.omit(Country)),
      studies       = max(studies, na.rm = TRUE),
      interventions = max(interventions, na.rm = TRUE),
      effective     = max(effective, na.rm = TRUE),
      .groups = "drop"
    )
  
  ws <- dplyr::left_join(world_sf, summary_df, by = ".cn")
  ws$Country <- dplyr::coalesce(ws$Country, ws$map_country)
  ws$Country <- tools::toTitleCase(tolower(ws$Country))
  
  for (nm in num_cols) {
    ws[[nm]] <- suppressWarnings(as.integer(as.numeric(ws[[nm]])))
    ws[[nm]][is.na(ws[[nm]])] <- 0L
  }
  
  ws <- tryCatch(sf::st_make_valid(ws), error = function(e) ws)
  keep <- tryCatch(!sf::st_is_empty(ws), error = function(e) rep(TRUE, nrow(ws)))
  ws <- ws[keep, , drop = FALSE]
  
  # Bin scheme: 0 | 1 | 2-5 | 6+
  vals_pos <- ws$studies[ws$studies >= 0]
  max_val <- if (length(vals_pos) > 0) max(vals_pos, na.rm = TRUE) else 0L
  
  if (max_val <= 0) {
    bins <- c(-0.5, 0.5)
    legend_labels <- c("0")
  } else if (max_val == 1) {
    bins <- c(-0.5, 0.5, 1.5)
    legend_labels <- c("0", "1")
  } else if (max_val <= 5) {
    bins <- c(-0.5, 0.5, 1.5, max_val + 0.5)
    legend_labels <- c("0", "1", "2+")
  } else {
    bins <- c(-0.5, 0.5, 1.5, 5.5, max_val + 0.5)
    legend_labels <- c("0", "1", "2\u20135", "6+")
  }
  
  # SFH palette, light to dark: Sand, then the Horizon Blue tones (the scale has
  # at most four classes: 0 | 1 | 2-5 | 6+).
  sfh_map_palette <- c(sfh_sand, sfh_blue_light, sfh_blue_mid, sfh_blue)
  
  n_intervals <- length(bins) - 1
  pal_cols <- sfh_map_palette[seq_len(min(length(sfh_map_palette), n_intervals))]
  if (length(pal_cols) < n_intervals) {
    pal_cols <- c(pal_cols, rep(sfh_ocean, n_intervals - length(pal_cols)))
  }
  
  pal <- colorBin(
    palette  = pal_cols,
    domain   = ws$studies,
    bins     = bins,
    na.color = sfh_sand,
    right    = TRUE
  )
  
  ws$fill_col <- ifelse(ws$studies > 0, pal(ws$studies), sfh_sand)
  ws$fill_op  <- ifelse(ws$studies > 0, 0.88, 0.30)
  
  labels <- sprintf(
    paste0(
      "<div style='font-family:General Sans, sans-serif; color:", sfh_charcoal, ";'>",
      "<strong>%s</strong><br/>",
      "Evaluations: %s<br/>",
      "Interventions: %s<br/>",
      "Reduced CSV evaluations: %s",
      "</div>"
    ),
    htmltools::htmlEscape(as.character(ws$Country)),
    ws$studies,
    ws$interventions,
    ws$effective
  ) %>% lapply(htmltools::HTML)
  
  leaflet(ws, options = leafletOptions(worldCopyJump = TRUE)) %>%
    addProviderTiles(providers$Esri.WorldGrayCanvas) %>%
    setView(lng = 20, lat = 20, zoom = zoom) %>%
    addPolygons(
      color        = "#FFFFFF",
      weight       = 0.8,
      opacity      = 1,
      smoothFactor = 0.4,
      fillColor    = ~fill_col,
      fillOpacity  = ~fill_op,
      label        = labels,
      labelOptions = labelOptions(
        direction = "auto",
        textsize  = "12px",
        style = list(
          "font-family"      = "General Sans, sans-serif",
          "font-weight"      = "600",
          "color"            = sfh_charcoal,
          "background-color" = "#FFFFFF",
          "border"           = "1px solid rgba(49,49,49,0.10)",
          "border-radius"    = "6px",
          "padding"          = "6px 8px"
        )
      ),
      highlight = highlightOptions(
        weight        = 2.2,
        color         = sfh_ocean,
        bringToFront  = TRUE,
        fillOpacity   = 0.95
      ),
      layerId = ~.cn
    ) %>%
    addLegend(
      position  = "bottomright",
      colors    = pal_cols,
      labels    = legend_labels,
      title     = paste0(
        "<div style='font-family:GT Pressura Extended, sans-serif; color:",
        sfh_charcoal, "; font-size:13px;'>",
        htmltools::htmlEscape(legend_title),
        "</div>"
      ),
      opacity   = 0.95,
      className = "sfh-legend",
      na.label  = "No data"
    )
}


# =============================================================================
# 10. UI definition
# =============================================================================
# Two tabs: Prevention explorer and Studies included.
# The CSS block below is kept verbatim - it controls the entire look and feel
# of the dashboard.

# One "label: value" line in the update-information box on the landing page:
# an orange bullet, a bold label, then the value.
update_info_line <- function(label, value) {
  tags$div(
    style = "display:inline-flex; align-items:center; gap:5px;",
    tags$span(
      "•",
      style = "color:#F09E56; font-weight:700; font-size:16px; line-height:1;"
    ),
    tags$span(
      tags$strong(style = "color:#313131;", label),
      value
    )
  )
}

ui <- navbarPage(
  title = "Dashboard of a living review of interventions to prevent childhood sexual violence victimisation and perpetration",
  theme = shinytheme("journal"),
  
  header = tagList(
    tags$head(
      tags$style(HTML("
  :root {
    /* Shared design tokens - keep card/panel treatments consistent site-wide */
    --sfh-card-radius: 14px;
    --sfh-card-shadow: 0 2px 10px rgba(0,0,0,0.04);
    --sfh-panel-radius: 12px;
    --sfh-space-sm: 8px;
    --sfh-space-md: 16px;
    --sfh-space-lg: 20px;
  }

  body {
  font-family: 'General Sans', sans-serif;
  background: linear-gradient(180deg, #FAFAF6 0%, #F3F2E8 100%);
  color: #313131;
  line-height: 1.55;
}

@font-face {
  font-family: 'General Sans';
  src: url('GeneralSans-Regular.otf') format('opentype');
}

@font-face {
  font-family: 'GT Pressura Extended';
  src: url('GT-Pressura-Extended-Regular.otf') format('opentype');
}

@font-face {
  font-family: 'DM Mono';
  src: url('DMMono-Regular.ttf') format('truetype');
}

.navbar {
  background: rgba(243, 242, 232, 0.96) !important;
  border-bottom: 1px solid rgba(49,49,49,0.12);
  box-shadow: 0 2px 10px rgba(0,0,0,0.03);
  backdrop-filter: blur(6px);
}

.navbar-brand {
  font-family: 'GT Pressura Extended', 'General Sans', sans-serif !important;
  font-weight: 700 !important;
  font-size: 20px !important;
  letter-spacing: 0.2px;
  color: #313131 !important;
  padding-top: 14px !important;
  padding-bottom: 14px !important;
  line-height: 1.2;
}

.navbar-brand:hover,
.navbar-brand:focus {
  color: #355571 !important;
}

.navbar-nav > li > a {
  font-family: 'General Sans', sans-serif !important;
  font-weight: 700 !important;
  font-size: 19px !important;
  color: #313131 !important;
  padding-top: 20px !important;
  padding-bottom: 20px !important;
  padding-left: 22px !important;
  padding-right: 22px !important;
  border-radius: 10px 10px 0 0;
  margin-top: 4px;
}

.navbar-nav > li > a:hover,
.navbar-nav > li > a:focus {
  color: #355571 !important;
  background: rgba(85,112,202,0.08) !important;
}

.navbar-nav > li.active > a,
.navbar-nav > li.active > a:hover,
.navbar-nav > li.active > a:focus {
  color: #355571 !important;
  background: #ffffff !important;
  border-top: 2px solid #F09E56 !important;
  box-shadow: 0 -1px 0 rgba(49,49,49,0.04);
}

h1 {
  font-family: 'GT Pressura Extended', 'General Sans', sans-serif;
  font-size: 2.0rem;
  line-height: 1.15;
  letter-spacing: 0.2px;
  margin-bottom: 0.75rem;
  color: #5570CA;
}

h2 {
  font-family: 'GT Pressura Extended', 'General Sans', sans-serif;
  font-size: 1.65rem;
  line-height: 1.2;
  margin-top: 1.35rem;
  margin-bottom: 0.75rem;
}

/* Title of the Prevention explorer and Studies included sections: the largest,
   boldest text in its block (the text underneath is set smaller, regular weight) */
h2.section-title {
  font-size: 2.4rem;
  font-weight: 700;
}

h3 {
  font-family: 'GT Pressura Extended', 'General Sans', sans-serif;
  font-size: 1.35rem;
  line-height: 1.22;
  margin-top: 1.15rem;
  margin-bottom: 0.7rem;
}

h4 {
  font-family: 'General Sans', sans-serif;
  font-weight: 700;
  font-size: 15px;
  letter-spacing: 0.1px;
  line-height: 1.25;
  margin-top: 1rem;
  margin-bottom: 0.6rem;
}

h5, h6 {
  font-family: 'General Sans', sans-serif;
  font-weight: 700;
  letter-spacing: 0.1px;
}
/* ===========================
   Header and footer logos
   =========================== */

/* Keep header title + logos aligned; let the title wrap onto a second
   line instead of being clipped when there isn't room for it */
.navbar {
  min-height: 64px;
  height: auto;
}

.navbar-header {
  min-height: 64px;
}

.navbar .container-fluid {
  padding-right: 270px;
  position: relative;
}

.navbar-brand {
  white-space: normal !important;
  overflow: visible;
  text-overflow: clip;
  max-width: calc(100vw - 520px);
  line-height: 1.25 !important;
  height: auto !important;
  padding-top: 12px !important;
  padding-bottom: 12px !important;
}

/* Top-right logos stay inside the navbar row */
.top-right-logos {
  position: absolute;
  top: 7px;
  right: 16px;
  z-index: 9999;
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 3px 6px;
  border-radius: 10px;
  background: rgba(243, 242, 232, 0.92);
  border: 1px solid rgba(49,49,49,0.08);
  pointer-events: none;
}

.top-right-logos img {
  height: 42px;
  width: auto;
  max-width: 118px;
  object-fit: contain;
  display: block;
}

/* Footer logos unchanged, but kept here for completeness */
.site-footer-logos {
  width: 100%;
  text-align: center;
  padding: 16px 0 12px 0;
  border-top: 1px solid rgba(49,49,49,0.12);
  margin-top: 24px;
  background: #F3F2E8;
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 28px;
  flex-wrap: wrap;
}

.site-footer-logos img {
  height: 52px;
  width: auto;
  max-width: 240px;
  object-fit: contain;
}

/* Medium screens: shrink title/logos but keep same row */
@media (max-width: 1100px) {
  .navbar .container-fluid {
    padding-right: 210px;
  }

  .navbar-brand {
    max-width: calc(100vw - 430px);
    font-size: 16px !important;
  }

  .top-right-logos {
    right: 12px;
    gap: 8px;
    padding: 2px 5px;
  }

  .top-right-logos img {
    height: 34px;
    max-width: 95px;
  }
}

/* Small screens: hide top logos instead of pushing them onto a new line */
@media (max-width: 760px) {
  .navbar {
    min-height: auto;
  }

  .navbar .container-fluid {
    padding-right: 15px;
  }

  .navbar-brand {
    max-width: calc(100vw - 90px);
    white-space: normal !important;
    font-size: 16px !important;
  }

  .top-right-logos {
    display: none;
  }

  .site-footer-logos img {
    height: 40px;
    max-width: 180px;
  }
}

table.dataTable tbody td {
  vertical-align: middle !important;
  line-height: 1.45 !important;
}

table.dataTable tbody td:nth-child(3),
table.dataTable tbody td:nth-child(4),
table.dataTable tbody td:nth-child(5) {
  min-width: 140px;
}

/* Indent rows inside groups */
table.dataTable tbody tr td:nth-child(4) {
  padding-left: 20px;
}

/* Remove indent from group headers */
tr.dtrg-group td {
  padding-left: 10px !important;
}

/* Make grouped sections clearer */
tr.dtrg-group {
  border-top: 2px solid #F09E56;
}

/* Slight spacing between groups */
tr.dtrg-group + tr td {
  padding-top: 6px;
}

@media (max-width: 600px) {
  .top-right-logos img {
    height: 30px;
    max-width: 105px;
  }

  .site-footer-logos img {
    height: 40px;
    max-width: 180px;
  }
}

.small-note, .dataTables_wrapper,
.modal-body, .panel-body {
  font-family: 'General Sans', sans-serif;
}

code, pre {
  font-family: 'DM Mono', monospace;
}

a { color: #4F9B90; }
a:hover { color: #355571; }

.welcome-box,
.home-widget,
.country-modal-card,
.panel.panel-default {
  background: #ffffff;
  border: 1px solid rgba(49,49,49,0.10);
  border-radius: var(--sfh-card-radius);
  box-shadow: var(--sfh-card-shadow);
}

.welcome-box {
  padding: var(--sfh-space-lg);
}

.home-widget {
  padding: var(--sfh-space-md);
}

.fluid-row,
.row {
  margin-bottom: 12px;
}

.btn,
.btn-default,
.btn-primary,
.action-button {
  background-color: #F09E56 !important;
  border-color: #C07E45 !important;
  color: #313131 !important;
  border-radius: 10px;
  font-weight: 700;
  font-family: 'General Sans', sans-serif;
  box-shadow: 0 1px 2px rgba(0,0,0,0.05);
}

.btn:hover,
.btn-default:hover,
.btn-primary:hover,
.action-button:hover {
  background-color: #F3B178 !important;
  border-color: #F3B178 !important;
  color: #1D1D1D !important;
  transform: translateY(-1px);
}

.btn:focus,
.btn-default:focus,
.btn-primary:focus,
.action-button:focus,
a:focus,
input:focus,
select:focus,
textarea:focus,
button:focus {
  outline: 3px solid rgba(85,112,202,0.35) !important;
  outline-offset: 2px !important;
  box-shadow: 0 0 0 3px rgba(85,112,202,0.15) !important;
}

.home-pill-row .pill-btn {
  background: #ffffff !important;
  color: #313131 !important;
  border: 1px solid rgba(49,49,49,0.18) !important;
  border-radius: 999px !important;
  padding: 9px 12px !important;
  font-size: 12px !important;
  line-height: 1.1 !important;
  white-space: nowrap;
  width: 100%;
}

.home-pill-row .pill-btn:hover {
  background: #F3F2E8 !important;
  color: #313131 !important;
}

.home-pill-row .pill-btn.pill-btn-primary {
  background: #4F9B90 !important;
  border: 1px solid #2F5D56 !important;
  color: #ffffff !important;
  font-weight: 700 !important;
  box-shadow: 0 2px 6px rgba(79,155,144,0.35) !important;
}

.home-pill-row .pill-btn.pill-btn-primary:hover {
  background: #2F5D56 !important;
  color: #ffffff !important;
}

.home-pill-row .pill-btn.pill-btn-primary .fa {
  margin-right: 6px;
}


.small-note,
.modal-body,
.panel-body,
.dataTables_wrapper {
  color: #313131;
}

table.dataTable thead th {
  font-family: 'GT Pressura Extended', 'General Sans', sans-serif;
  font-weight: 600;
  font-size: 13px;
  letter-spacing: 0.4px;
  text-transform: none;
  color: #1D1D1D !important;
  background: #F3F2E8 !important;
  border-bottom: 2px solid rgba(49,49,49,0.15) !important;
}

table.dataTable tbody td {
  font-family: 'General Sans', sans-serif;
  font-size: 13px;
  color: #313131 !important;
  padding: 10px 12px;
}

table.dataTable tbody tr:hover {
  background-color: #FAFAF6 !important;
}

table.dataTable tbody tr:nth-child(even) {
  background-color: #FAFAF6;
}

.dataTables_wrapper,
.dataTables_wrapper * {
  font-family: 'General Sans', sans-serif !important;
}

.dataTables_wrapper .dataTables_paginate .paginate_button {
  font-family: 'General Sans', sans-serif;
  color: #313131 !important;
  border: 1px solid transparent !important;
  border-radius: 6px;
  background: transparent !important;
  box-shadow: none !important;
}

.dataTables_wrapper .dataTables_paginate .paginate_button:hover {
  background: #F3F2E8 !important;
  border: 1px solid rgba(49,49,49,0.15) !important;
  color: #313131 !important;
}

.dataTables_wrapper .dataTables_paginate .paginate_button.current,
.dataTables_wrapper .dataTables_paginate .paginate_button.current:hover {
  background: #F09E56 !important;
  border: 1px solid #C07E45 !important;
  color: #fff !important;
  border-radius: 6px;
}

.dataTables_wrapper .dataTables_paginate .paginate_button.disabled,
.dataTables_wrapper .dataTables_paginate .paginate_button.disabled:hover {
  background: transparent !important;
  border-color: transparent !important;
  color: rgba(49,49,49,0.35) !important;
}

.dataTables_wrapper .dataTables_filter input,
.dataTables_wrapper .dataTables_length select {
  font-family: 'General Sans', sans-serif;
  border-radius: 6px;
  border: 1px solid rgba(49,49,49,0.2);
  background: #ffffff;
  color: #313131;
  padding: 4px 8px;
}

.dataTables_wrapper .dataTables_info,
.dataTables_wrapper .dataTables_length,
.dataTables_wrapper .dataTables_filter label {
  color: #313131 !important;
}

details summary {
  list-style: none;
  cursor: pointer;
  font-family: 'GT Pressura Extended', 'General Sans', sans-serif;
  font-weight: 700;
  color: #355571;
}

details summary::-webkit-details-marker {
  display: none;
}

details summary::before {
  content: '▸ ';
  color: #F09E56;
  font-weight: 700;
}

details[open] summary::before {
  content: '▾ ';
}

.sfh-legend {
  background: rgba(255,255,255,0.96) !important;
  border: 1px solid rgba(49,49,49,0.12) !important;
  border-radius: var(--sfh-panel-radius) !important;
  box-shadow: 0 2px 8px rgba(0,0,0,0.08) !important;
  padding: 10px 12px !important;
  color: #313131 !important;
  line-height: 1.35 !important;
  font-family: 'General Sans', sans-serif !important;
}

.sfh-legend .legend-title {
  font-family: 'GT Pressura Extended', 'General Sans', sans-serif !important;
  font-size: 13px !important;
  font-weight: 600 !important;
  color: #313131 !important;
  margin-bottom: 6px !important;
}

.sfh-legend .legend-labels,
.sfh-legend .legend-labels span,
.sfh-legend .legend-labels div {
  font-family: 'General Sans', sans-serif !important;
  font-size: 12px !important;
  color: #313131 !important;
}

.sfh-legend i {
  width: 16px !important;
  height: 16px !important;
  margin-right: 8px !important;
  margin-top: 1px !important;
  opacity: 1 !important;
  border: 1px solid rgba(49,49,49,0.10) !important;
}

.leaflet-tooltip,
.leaflet-tooltip p,
.leaflet-tooltip strong,
.leaflet-popup-content,
.leaflet-popup-content p,
.leaflet-control,
.leaflet-control .info,
.leaflet-control .legend,
.leaflet-control .legend-labels {
  font-family: 'General Sans', sans-serif !important;
  font-size: 12px !important;
  color: #313131 !important;
}

.country-modal-card {
  padding: 14px 16px;
  height: 100%;
}

.country-modal-card .country-modal-label {
  font-family: 'General Sans', sans-serif !important;
  font-size: 12px;
  color: #5A5A5A;
  margin-bottom: 6px;
}

.country-modal-card .country-modal-value {
  font-family: 'GT Pressura Extended', 'General Sans', sans-serif;
  font-size: 28px;
  line-height: 1.05;
  font-weight: 700;
  color: #313131;
}

.country-modal-card .country-modal-accent {
  height: 4px;
  border-radius: 999px;
  margin-bottom: 10px;
}

.country-modal-summary {
  font-family: 'GT Pressura Extended', 'General Sans', sans-serif !important;
  font-size: 14px;
  color: #355571;
  cursor: pointer;
}

.country-modal-table {
  width: 100%;
  border-collapse: collapse;
  font-family: 'General Sans', sans-serif !important;
  margin-bottom: 0;
}

.country-modal-table th {
  background: #F3F2E8;
  border-bottom: 2px solid rgba(49,49,49,0.12);
  padding: 8px 10px;
  font-family: 'GT Pressura Extended', 'General Sans', sans-serif !important;
  font-size: 12px;
  color: #313131;
  text-align: left;
}

.country-modal-table td {
  padding: 8px 10px;
  border-bottom: 1px solid rgba(49,49,49,0.08);
  font-size: 13px;
  vertical-align: top;
  color: #313131;
}

.country-modal-table tr:nth-child(even) {
  background: #FAFAF6;
}

@media (max-width: 768px) {
  .navbar-brand {
    font-size: 16px !important;
  }

  h1 { font-size: 1.5rem; }
  h2 { font-size: 1.3rem; }
  h2.section-title { font-size: 1.8rem; }
  h3 { font-size: 1.2rem; }

  .home-pill-row {
    grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)) !important;
  }
}

/* Prevention explorer table polish */
#intervention_table table.dataTable thead th,
#intervention_table table.dataTable tbody td {
  text-align: center !important;
}

#intervention_table table.dataTable tbody td {
  vertical-align: middle !important;
  white-space: normal !important;
  word-break: break-word !important;
  line-height: 1.35 !important;
}


#intervention_table span {
font-size: 11px !important;
line-height: 1.2 !important;
  white-space: normal !important;
}

/* Column order: 1 New, 2 View, 3 Intervention name, 4 Region, 5 Country,
   6 Type of CSV outcome(s), 7 Included age groups, 8 Location of
   implementation, 9 INSPIRE category, 10 # Evaluations since 2018,
   11 Public access */
#intervention_table table.dataTable tbody td:nth-child(3) {
  min-width: 190px !important;
  white-space: normal !important;
}
#intervention_table table.dataTable tbody td:nth-child(4) {
  min-width: 130px !important;
  white-space: normal !important;
}
#intervention_table table.dataTable tbody td:nth-child(5) {
  min-width: 140px !important;
  white-space: normal !important;
}
#intervention_table table.dataTable tbody td:nth-child(6) {
  min-width: 190px !important;
  white-space: normal !important;
}
#intervention_table table.dataTable tbody td:nth-child(7) {
  min-width: 190px !important;
  white-space: normal !important;
}
#intervention_table table.dataTable tbody td:nth-child(8) {
  min-width: 150px !important;
  white-space: normal !important;
}
#intervention_table table.dataTable tbody td:nth-child(9) {
  min-width: 170px !important;
  white-space: normal !important;
}
#intervention_table table.dataTable tbody td:nth-child(10) {
  min-width: 110px !important;
  white-space: nowrap !important;
}
#intervention_table table.dataTable tbody td:nth-child(11) {
  min-width: 170px !important;
  white-space: normal !important;
}

#intervention_table div::-webkit-scrollbar {
  width: 6px;
}

#intervention_table div::-webkit-scrollbar-thumb {
  background: #C2C2BA;
  border-radius: 6px;
}

")),

      # Intro panel toggle. Nothing is stored in the visitor's browser: the panel
      # always starts expanded and visitors can collapse it themselves.
      tags$script(HTML("
    (function() {
      document.addEventListener('DOMContentLoaded', function() {
        var btn = document.getElementById('sfh-intro-toggle');
        if (!btn) return;
        function sync() {
          var collapsed = document.documentElement.classList.contains('sfh-intro-collapsed');
          btn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
        }
        sync();
        btn.addEventListener('click', function() {
          document.documentElement.classList.toggle('sfh-intro-collapsed');
          sync();
        });
      });
    })();
  "))
    ),
    tags$div(
      class = "top-right-logos",
      tags$img(src = "sfh_full_colour_centralised_rgb.png", alt = "Safe Futures Hub logo"),
      tags$img(src = "DSPI_Primary_logo.jpg", alt = "Department of Social Policy and Intervention logo")
    )
  ),
  
  # ---------------------------------------------------------------------------
  # Tab 1: Prevention explorer
  # ---------------------------------------------------------------------------
  tabPanel(
    "Prevention explorer",
    fluidPage(
      tags$head(
        tags$style(HTML(paste0(
          "body { background: ", bg_white, "; }",
          ".small-note { font-size:12px; color:#5A5A5A; margin-top:6px; }",
          ".dataTables_wrapper table.dataTable td { white-space: normal; }"
        ))),
        tags$style(HTML("
  /* Hover tooltip for study design definitions */
  .sd-tooltip {
    position: relative;
    cursor: help;
    border-bottom: 1px dotted #355571;
  }

  .sd-tooltip::after {
    content: attr(data-tooltip);
    position: absolute;
    left: 0;
    bottom: 130%;
    background: #313131;
    color: #fff;
    padding: 8px 10px;
    border-radius: 8px;
    font-size: 12px;
    font-weight: 400;
    line-height: 1.4;
    white-space: normal;
    width: max-content;
    max-width: 280px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.18);
    opacity: 0;
    visibility: hidden;
    transform: translateY(4px);
    transition: opacity 0.15s ease, transform 0.15s ease;
    z-index: 2000;
    pointer-events: none;
  }

  .sd-tooltip::before {
    content: '';
    position: absolute;
    left: 12px;
    bottom: 122%;
    border: 6px solid transparent;
    border-top-color: #313131;
    opacity: 0;
    visibility: hidden;
    transition: opacity 0.15s ease;
    z-index: 2000;
  }

  .sd-tooltip:hover::after,
  .sd-tooltip:hover::before {
    opacity: 1;
    visibility: visible;
  }
")),
        tags$style(HTML("
  /* Forward-moving staggered reveal, played on page load and on tab switch */
  @keyframes sfh-forward-reveal {
    from { opacity: 0; transform: translateX(-16px); }
    to   { opacity: 1; transform: translateX(0); }
  }

  .sfh-stagger.sfh-stagger-init > * {
    opacity: 0;
  }

  .sfh-stagger.sfh-stagger-play > * {
    animation: sfh-forward-reveal 0.5s ease-out forwards;
  }

  .sfh-stagger.sfh-stagger-play > *:nth-child(1) { animation-delay: 0s; }
  .sfh-stagger.sfh-stagger-play > *:nth-child(2) { animation-delay: 0.1s; }
  .sfh-stagger.sfh-stagger-play > *:nth-child(3) { animation-delay: 0.2s; }
  .sfh-stagger.sfh-stagger-play > *:nth-child(n+4) { animation-delay: 0.3s; }

  @media (prefers-reduced-motion: reduce) {
    .sfh-stagger > * {
      opacity: 1 !important;
      animation: none !important;
      transform: none !important;
    }
  }
")),
        # Styling for the always-visible "More information" section
        tags$style(HTML("
  .dashboard-info-details {
    flex-basis: 100%;
    width: 100%;
  }

  .dashboard-info-details .dashboard-info-expanded {
    width: 100%;
    box-sizing: border-box;
  }

  /* Landing layout: text column + compact rail (dates and action buttons) */
  .sfh-landing-grid {
    display: grid;
    grid-template-columns: minmax(0, 1fr) 300px;
    gap: 20px;
    align-items: start;
    margin-top: 10px;
    margin-bottom: 4px;
  }

  .sfh-landing-rail {
    min-width: 0;
  }

  @media (max-width: 900px) {
    .sfh-landing-grid {
      grid-template-columns: minmax(0, 1fr);
    }
  }

  /* Collapsible intro toggle (button-style, same look as the old label) */
  .sfh-intro-toggle {
    display: inline-flex;
    align-items: center;
    gap: 10px;
    width: fit-content;
    max-width: 100%;
    font-family: 'GT Pressura Extended', 'General Sans', sans-serif;
    font-size: 15px;
    font-weight: 700;
    text-align: left;
    color: #313131;
    padding: 10px 14px;
    background: #F09E56;
    border: 1px solid #C07E45;
    border-radius: 12px;
    box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    cursor: pointer;
    transition: background-color 0.15s ease, transform 0.15s ease;
  }

  .sfh-intro-toggle:hover {
    background: #F3B178;
    transform: translateY(-1px);
  }

  /* Decorative chevron: points down when open, right when collapsed */
  .sfh-intro-toggle::after {
    content: '';
    flex: none;
    width: 8px;
    height: 8px;
    border-right: 2px solid #313131;
    border-bottom: 2px solid #313131;
    transform: translateY(-2px) rotate(45deg);
    transition: transform 0.2s ease;
  }

  html.sfh-intro-collapsed .sfh-intro-toggle::after {
    transform: translateX(-2px) rotate(-45deg);
  }

  html.sfh-intro-collapsed .dashboard-info-expanded {
    display: none;
  }

  @media (prefers-reduced-motion: reduce) {
    .sfh-intro-toggle,
    .sfh-intro-toggle::after {
      transition: none;
    }
    .sfh-intro-toggle:hover { transform: none; }
  }

  /* Clickable navigation cards inside the info panel */
  .sfh-nav-cards {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
    gap: 10px;
    margin: 0 0 10px 0;
  }

  a.sfh-nav-card {
    position: relative;
    display: block;
    padding: 12px 34px 12px 14px;
    background: #ffffff;
    border: 1px solid rgba(49,49,49,0.12);
    border-radius: var(--sfh-panel-radius);
    color: #313131;
    text-decoration: none;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    transition: border-color 0.15s ease, box-shadow 0.15s ease, transform 0.15s ease;
  }

  a.sfh-nav-card strong {
    color: #355571;
  }

  /* Decorative arrow (CSS only, so the text itself is unchanged) */
  a.sfh-nav-card::after {
    content: '\\2192';
    position: absolute;
    right: 12px;
    top: 10px;
    font-size: 18px;
    line-height: 1;
    color: #F09E56;
    font-weight: 700;
    transition: transform 0.15s ease;
  }

  a.sfh-nav-card:hover,
  a.sfh-nav-card:focus {
    color: #313131;
    text-decoration: none;
    border-color: #F09E56;
    box-shadow: 0 3px 10px rgba(0,0,0,0.08);
    transform: translateY(-1px);
  }

  a.sfh-nav-card:hover::after,
  a.sfh-nav-card:focus::after {
    transform: translateX(3px);
  }

  @media (prefers-reduced-motion: reduce) {
    a.sfh-nav-card,
    a.sfh-nav-card::after {
      transition: none;
      transform: none !important;
    }
  }
")),
        
        
        tags$script(HTML(js)),
        tags$script(HTML("
    $(document).on('click', '.filter-chip', function() {
      var value = $(this).data('value');
      var inputId = $(this).data('input');

      Shiny.setInputValue('remove_filter', {
        value: value,
        input: inputId
      }, {priority: 'event'});
    });
  ")),
        tags$script(HTML("
    // Landing-page navigation cards: 'tab:<name>' switches navbar tab,
    // 'section:<id>' scrolls to an element on the current tab.
    $(document).on('click', 'a.sfh-nav-card', function(e) {
      e.preventDefault();
      var target = String($(this).data('nav-target') || '');
      var idx = target.indexOf(':');
      var kind = target.slice(0, idx);
      var value = target.slice(idx + 1);

      if (kind === 'tab') {
        var link = $('.navbar-nav a').filter(function() {
          return $(this).data('value') === value;
        }).first();
        if (link.length) {
          link.trigger('click');
          window.scrollTo({ top: 0, behavior: 'smooth' });
        }
      } else if (kind === 'section') {
        var el = document.getElementById(value);
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });
  ")),
        tags$script(HTML("
    function sfhPlayStagger(root) {
      var scope = root && root.querySelectorAll ? root : document;
      var containers = scope.querySelectorAll('.sfh-stagger');
      containers.forEach(function(el) {
        el.classList.remove('sfh-stagger-play');
        el.classList.add('sfh-stagger-init');
        requestAnimationFrame(function() {
          requestAnimationFrame(function() {
            el.classList.remove('sfh-stagger-init');
            el.classList.add('sfh-stagger-play');
          });
        });
      });
    }

    $(document).ready(function() {
      sfhPlayStagger(document);
    });

    $(document).on('shown.bs.tab', 'a[data-toggle=\"tab\"]', function(e) {
      var target = $(e.target).attr('href');
      var pane = target ? document.querySelector(target) : null;
      sfhPlayStagger(pane || document);
    });
  "))
      ),
      
      fluidRow(
        column(
          12,
          div(
            class = "welcome-box",
            fluidRow(
              column(
                12,
                tags$div(
                  class = "sfh-stagger",
                  h2("Welcome to this dashboard on global evidence to prevent childhood sexual violence victimisation and perpetration"),
                  
                  # Two-column landing layout: long text on the left, a compact
                  # rail (update dates + action buttons) on the right. Stacks
                  # to a single column on narrow screens (see CSS above).
                  tags$div(
                    class = "sfh-landing-grid",

                    tags$div(
                    class = "sfh-landing-main",

                    # More information section (always visible, no longer a click-to-expand toggle)
                    tags$div(
                      class = "dashboard-info-details",
                      style = "margin:0;",

                      # Toggle button: the panel starts expanded; clicking collapses or
                      # expands it (see the small script in the page header).
                      tags$button(
                        type = "button",
                        id = "sfh-intro-toggle",
                        class = "sfh-intro-toggle",
                        `aria-expanded` = "true",
                        `aria-controls` = "sfh-intro-body",
                        "More information about this dashboard"
                      ),

                      tags$div(
                        id = "sfh-intro-body",
                        class = "dashboard-info-expanded",
                        style = "
        margin-top:10px;
        padding:14px;
        background:#F3F2E8;
        border-radius:var(--sfh-panel-radius);
        border-left:4px solid #F09E56;
      ",

                        tags$h4(
                          style = "
          font-family:'GT Pressura Extended', sans-serif;
          font-size:19px;
          font-weight:700;
          color:#355571;
          margin-top:0;
        ",
                          "What is included in this dashboard?"
                        ),

                        tags$p(
                          "Research on preventing childhood sexual violence (CSV) victimisation and perpetration is constantly evolving. This dashboard presents the latest evidence."
                        ),

                        tags$p(
                          "This dashboard was developed by the ",
                          tags$a(href = "https://www.safefutureshub.org/", target = "_blank", "Safe Futures Hub"),
                          " and the ",
                          tags$a(href = "https://www.spi.ox.ac.uk/home", target = "_blank", "Department of Social Policy and Intervention at the University of Oxford"),
                          ". It shares up-to-date evidence on evaluations of interventions, laws and policies to prevent childhood sexual violence victimisation (e.g., sexual violence, abuse, harassment) and perpetration (e.g., harmful sexual behaviour against children from adults or other children) published since 2018."
                        ),

                        tags$p(
                          "The dashboard is based on a living systematic review. It is periodically updated with peer-reviewed research from around the world, ensuring the evidence reflects the latest findings as new studies are published."
                        ),

                        tags$p(
                          "You can explore where interventions have been implemented and evaluated, who they target, and the CSV victimisation and perpetration outcomes they aim to prevent. You can filter by region, country, age group, outcome, and context of intervention to identify interventions relevant to your setting."
                        ),

                        # Clickable cards: each one jumps to the part of the dashboard
                        # it describes (handled by the navigation-card script in the head).
                        tags$div(
                          class = "sfh-nav-cards",
                          tags$a(
                            href = "#",
                            class = "sfh-nav-card",
                            `data-nav-target` = "section:prevention-explorer-section",
                            "Use the ",
                            tags$strong("Prevention explorer"),
                            " to see what is effective. It highlights interventions, laws, and policies shown to reduce childhood sexual violence victimisation or perpetration since 2018."
                          ),
                          tags$a(
                            href = "#",
                            class = "sfh-nav-card",
                            `data-nav-target` = "tab:Studies included",
                            "Use the ",
                            tags$strong("Studies included"),
                            " to see all reviewed evidence. It shares the full body of peer-reviewed and published research, including studies that reduced, increased, and had no effects on childhood sexual violence victimisation and perpetration."
                          )
                        ),

                        tags$p(
                          "This dashboard informs the intervention case studies available on the Safe Futures Hub. For practical summaries and implementation insights on specific interventions, visit ",
                          tags$a(href = "https://www.safefutureshub.org/solutions", target = "_blank", "https://www.safefutureshub.org/solutions"),
                          "."
                        )
                      )
                    )
                    ),

                    tags$div(
                    class = "sfh-landing-rail",

                    # Compact update information
                    tags$div(
                      style = "
      display:flex;
      flex-direction:column;
      width:100%;
      box-sizing:border-box;
      max-width:100%;
      gap:6px;
      padding:9px 12px;
      background:#FAFAF6;
      border:1px solid rgba(49,49,49,0.10);
      border-radius:10px;
      font-family:'General Sans', sans-serif;
      font-size:12px;
      color:#5A5A5A;
    ",

                      update_info_line("Evidence current as of:", " 9 Dec 2025"),
                      update_info_line("Dashboard last updated:", " 1 October 2026"),
                      update_info_line("Next dashboard update anticipated:", " 2027")
                    ),

                    tags$div(
                    class = "home-pill-row",
                    style = "
              display: grid;
              grid-template-columns: repeat(2, minmax(0, 1fr));
              gap: 10px;
              margin-top: 12px;
            ",
                    actionButton("open_dashboard_help", "How to use", icon = icon("circle-question"), class = "btn btn-default pill-btn pill-btn-primary"),
                    actionButton("open_approach_modal", "Approach", icon = icon("circle-info"), class = "btn btn-default pill-btn"),
                    actionButton("view_definitions_home", "Definitions", class = "btn btn-default pill-btn"),
                    actionButton("reset_filters_home", "Reset all filters", class = "btn btn-default pill-btn"),
                    actionButton("download_summary_all", "Download data", class = "btn btn-default pill-btn"),
                    actionButton("cite_dashboard", "Cite Dashboard", class = "btn btn-default pill-btn"),
                    actionButton("open_feedback_form", "Give feedback", class = "btn btn-default pill-btn"),
                    actionButton("open_share_evidence_form", "Share evidence", class = "btn btn-default pill-btn")
                    )
                    )
                  )
                )
              )
            )
          )
        )
      ),
      br(),
      
      fluidRow(
        column(
          12,
          h2(id = "prevention-explorer-section", class = "section-title", "Prevention explorer: A summary of interventions that reduced childhood sexual violence (CSV) victimisation and perpetration"),
          uiOutput("status_line_intervention_static"),
          
          tags$div(
            class = "welcome-box",
            style = "margin-bottom:14px;",
            
            tags$div(
              style = "
    font-family: 'GT Pressura Extended', sans-serif;
    font-size: 16px;
    font-weight: 700;
    color: #355571;
    margin-bottom: 6px;
  ",
              "Explore interventions that reduced CSV victimisation or perpetration"
            ),

            tags$p(
              style = "margin-bottom:10px; margin-top:0; color:#313131;",
              "Use some or all of the filters below to explore where interventions that reduced CSV victimisation or perpetration have been implemented, what outcomes were addressed, and which age groups were targeted."
            ),
            
            fluidRow(
              column(12, class = "col-md-3",
                     tags$label(
                       style = "font-family:'GT Pressura Extended', sans-serif; color:#355571; margin-bottom:6px;",
                       "Where is the intervention?"
                     ),
                     selectizeInput(
                       "int_region_filter",
                       "Region",
                       choices = sort(unique(na.omit(df$World.Region))),
                       selected = NULL,
                       multiple = TRUE
                     ),
                     selectizeInput(
                       "int_country_filter",
                       "Country",
                       choices = country_choices,
                       selected = NULL,
                       multiple = TRUE
                     )
              ),
              
              column(12, class = "col-md-3",
                     tags$label(
                       style = "font-family:'GT Pressura Extended', sans-serif; color:#355571; margin-bottom:6px;",
                       "What type of CSV is addressed?"
                     ),
                     selectizeInput(
                       "int_outcome_filter",
                       "Type of CSV victimisation or perpetration outcome(s)",
                       choices = sort(unique(unlist(df$Outcome_Domain_List))),
                       selected = NULL,
                       multiple = TRUE
                     )
              ),
              
              column(12, class = "col-md-3",
                     tags$label(
                       style = "font-family:'GT Pressura Extended', sans-serif; color:#355571; margin-bottom:6px;",
                       "Who is the intervention designed for?"
                     ),
                     selectizeInput(
                       "int_age_filter",
                       "Age groups",
                       choices = order_age_groups(unlist(df$Age.Groups_List)),
                       selected = NULL,
                       multiple = TRUE
                     )
              ),
              
              column(12, class = "col-md-3",
                     tags$label(
                       style = "font-family:'GT Pressura Extended', sans-serif; color:#355571; margin-bottom:6px;",
                       "How is the intervention delivered?"
                     ),
                     selectizeInput(
                       "int_level_filter",
                       "Location of implementation",
                       choices = sort(unique(unlist(df$Level_List))),
                       selected = NULL,
                       multiple = TRUE
                     ),
                     selectizeInput(
                       "int_inspire_filter",
                       "INSPIRE strategy",
                       choices = inspire_choices,
                       selected = NULL,
                       multiple = TRUE
                     )
              )
            ),
            
            tags$div(
              style = "margin-top:6px; display:flex; gap:8px; flex-wrap:wrap;",
              actionButton("int_quick_school", "School", class = "btn btn-sm btn-default"),
              actionButton("int_quick_community", "Community", class = "btn btn-sm btn-default"),
              actionButton("int_quick_adolescents", "Adolescents", class = "btn btn-sm btn-default"),
              actionButton("int_quick_adults", "Adults", class = "btn btn-sm btn-default"),
              actionButton("int_quick_reset", "Reset", class = "btn btn-sm btn-default")
            ),
            
            tags$div(
              style = "margin-top:6px; display:flex; gap:8px; flex-wrap:wrap;",
              lapply(sort(unique(na.omit(df$World.Region))), function(rg) {
                actionButton(
                  paste0("int_quick_region_", make.names(rg)),
                  rg,
                  class = "btn btn-sm btn-default"
                )
              })
            ),
            
            tags$p(
              style = "margin-top:10px; margin-bottom:0; font-size:12px; color:#5A5A5A;",
              tags$strong("Tip: "),
              "Start with region or country, then narrow by what the intervention addresses."
            )
          ),
          
          uiOutput("status_line_intervention_explorer"),
          br(),
          
          DTOutput("intervention_table"),
          br(),
          tags$p(
            "Please let us know if you plan to implement any interventions from this dashboard at ",
            tags$a(href = "mailto:info@safefutureshub.org", "info@safefutureshub.org"),
            "."
          ),
          downloadButton("download_intervention_summary", "Download intervention summary")
        )
      )
    )
  ),
  
  # ---------------------------------------------------------------------------
  # Tab 2: Studies included
  # ---------------------------------------------------------------------------
  tabPanel(
    "Studies included",
    fluidPage(

      # --- Static summary graphs ---
      tags$div(
        class = "sfh-stagger",
        h2("Studies included: An overview of academic evidence to prevent childhood sexual violence (CSV) from 2018."),
        uiOutput("status_line_evidence_static")
      ),
      fluidRow(
        column(6, tags$div(class="home-widget", h5("Where are prevention efforts located?"),
                           plotOutput("counts_bar", height="260px") %>% withSpinner())),
        column(6, tags$div(class="home-widget", h5("Which age groups are included?"),
                           plotOutput("age_plotly", height="260px") %>% withSpinner()))
      ),
      fluidRow(
        column(6, tags$div(class="home-widget", h5("Which countries is the evidence from?"),
                           plotOutput("income_plotly", height="260px") %>% withSpinner())),
        column(6, tags$div(class="home-widget", h5("How has the evidence base developed over time?"),
                           plotOutput("timeline_plotly", height="260px") %>% withSpinner()))
      ),
      
      br(),
      fluidRow(
        column(
          12,
          h2("Interactive world map of prevention effort to address childhood sexual violence victimisation and perpetration since 2018"),
          p(class = "small-note",
            "Countries with evaluations of child sexual violence prevention (CSV) efforts published after 2018 where CSV victimisation and perpetration were measured are highlighted (see legend for more information).",
            "Hover over a country to find the total number of evaluations, interventions, and effective evaluations for that country."
          ),
          leafletOutput("world_map_home", height = "50vh") %>% withSpinner()
        )
      ),
      hr(),
      
      # --- Interactive explorer ---
      sidebarLayout(
        sidebarPanel(
          width = 3,
          h4("Filters for studies included"),
          
          h5("Context & geography"),
          selectizeInput("country_filter", "Country",
                         choices = country_choices, selected = NULL, multiple = TRUE),
          selectizeInput("region_filter", "Region",
                         choices = sort(unique(na.omit(df$World.Region))), selected = NULL, multiple = TRUE),
          selectizeInput("income_filter", "Income classification",
                         choices = sort(unique(na.omit(df$Income.Classification))), selected = NULL, multiple = TRUE),
          
          h5("CSV outcomes"),
          selectizeInput("outcome_domain_filter", "CSV perpetration or victimisation",
                         choices = outcome_domain_choices, selected = NULL, multiple = TRUE),
          selectizeInput("specific_outcome_filter", "Specific victimisation or perpetration outcome",
                         choices = specific_outcome_choices,
                         selected = NULL, multiple = TRUE),
          
          h5("Evidence rating"),
          selectizeInput("evidence_rating_filter", "Evidence rating",
                         choices = sort(unique(na.omit(df$evidence_rating))), selected = NULL, multiple = TRUE),
          
          h5("Study populations & context of intervention"),
          selectizeInput("population_filter", "Age Group",
                         choices = order_age_groups(unlist(df$Age.Groups_List)),
                         selected = NULL, multiple = TRUE),
          selectizeInput("context_by_filter", "Location of intervention",
                         choices = context_by_filter,
                         selected = NULL, multiple = TRUE),
          selectizeInput("inspire_filter", "INSPIRE strategies",
                         choices = inspire_choices,
                         selected = NULL, multiple = TRUE),
          hr(),
          h5("Quick views"),
          actionButton("quick_adolescents", "Adolescents (age 10-18 years)", class = "btn btn-default btn-sm"),
          br(), br(),
          actionButton("quick_school", "Location: School-based", class = "btn btn-default btn-sm"),
          br(), br(),
          actionButton("quick_community", "Location: Community-based", class = "btn btn-default btn-sm"),
          br(), br(),
          actionButton("quick_online", "Location: Online", class = "btn btn-default btn-sm"),
          br(), br(),
          hr(),
          actionButton("reset_filters", "Reset all filters", class = "btn btn-default")
        ),
        
        mainPanel(
          width = 9,
          h2(class = "section-title", "Interactive explorer for all included studies"),

          tags$div(
            style = "margin-top:4px;",
            tags$p(
              style = "margin-top:6px; margin-bottom:0; color:#313131;",
              "The studies listed in the table on this tab include all evaluations of interventions published since 2018 that either found a reduction, no change, or an increase in childhood sexual violence (CSV) victimisation or perpetration outcomes."
            )
          ),

          tags$div(
            class = "home-status-line",
            uiOutput("status_line_explorer"),
            uiOutput("active_filter_chips")
          ),
          
          tags$div(
            style = "margin:10px 0 14px 0;",
            tags$div(
              style = "font-family:'GT Pressura Extended', sans-serif; color:#355571; font-size:14px;",
              "Definitions of evidence ratings"
            ),
            tags$div(
              style = "margin-top:10px; padding:12px; background:#F3F2E8; border:1px solid rgba(49,49,49,0.10); border-radius:var(--sfh-panel-radius);",
              tags$p(
                "We examined the authors’ reports on the effectiveness of the intervention. We assessed all CSV victimisation and perpetration outcomes reported in the studies and grouped them into three groups:"
              ),
              tags$p(
                tags$strong("Reduced childhood sexual violence (CSV):"),
                " These are evaluations which found a reduction in ",
                tags$strong("at least one"),
                " CSV victimisation or perpetration outcome and no increase in ",
                tags$strong("any other"),
                " CSV victimisation or perpetration outcomes."
              ),
              tags$p(
                tags$strong("No change in childhood sexual violence (CSV):"),
                " These are evaluations which found no change in ",
                tags$strong("all"),
                " CSV victimisation or perpetration outcomes."
              ),
              tags$p(
                tags$strong("Increased childhood sexual violence (CSV):"),
                " These are evaluations which found an increase in ",
                tags$strong("at least one"),
                " CSV victimisation or perpetration outcome."
              )
            )
          ),
          
          tags$div(
            style = "margin-bottom:12px;",
            actionButton("open_dashboard_help_explorer", "How to use this dashboard", class = "btn btn-default")
          ),
          
          DTOutput("results_dt") %>% withSpinner(),
          br(),
          downloadButton("download_all", "Download included studies summary")
        )
      )
    )
  ),
  
  # ---------------------------------------------------------------------------
  # Footer
  # ---------------------------------------------------------------------------
  footer = tags$div(
    class = "site-footer-logos",
    tags$img(src = "sfh_full_colour_centralised_rgb.png", alt = "Safe Futures Hub logo"),
    tags$img(src = "DSPI_Primary_logo_Ox.jpg", alt = "Department of Social Policy and Intervention and University of Oxford logo")
  )
)


# =============================================================================
# 10b. Study design definitions (for hover tooltips)
# =============================================================================
study_design_definitions <- c(
  "RCT" = "Randomised controlled trial: individuals are randomly assigned to intervention or control groups to estimate the causal effect of an intervention.",
  "Cluster RCT" = "Cluster randomised controlled trial: groups (e.g., schools, hospitals, communities) rather than individuals are randomly assigned to intervention or control.",
  "Pilot trial" = "A small-scale preliminary study conducted to assess the feasibility, acceptability, and procedures for a future definitive trial.",
  "Non-randomised controlled trial" = "Participants are allocated to intervention and control groups without randomisation, often based on availability or preference.",
  "Difference-in-differences" = "A quasi-experimental design that estimates an intervention effect by comparing changes over time between intervention and comparison groups.",
  "Interrupted time series" = "Evaluates the impact of an intervention by examining changes in outcome level and/or trend before and after its implementation using repeated observations over time.",
  "Propensity score matching/weighted comparison" = "A quasi-experimental approach that uses propensity scores to create intervention and comparison groups with similar measured baseline characteristics.",
  "Regression discontinuity" = "A quasi-experimental design that estimates the intervention effect by comparing individuals just above and below a predefined eligibility threshold.",
  "Pre-post (single group, no comparison)" = "Outcomes are measured before and after an intervention in the same group, with no control or comparison group.",
  "Longitudinal cohort" = "A cohort of individuals is followed over time to assess changes in outcomes and/or associations with exposures.",
  "Cross-sectional comparison" = "Exposure and outcome are measured at a single point in time, allowing comparison between groups but not assessment of temporal relationships.",
  "Quasi-experimental study" = "A broad category of intervention studies that estimate causal effects without randomisation, often using natural experiments or statistical methods to address confounding.",
  "Other" = "A study design that does not fit into the categories above."
)

# Renders a study design value as a hoverable span carrying its definition.
# Classifies the first line of the raw value by keyword (case-insensitive),
# not exact/prefix matching against the canonical terms above - the source
# data spells designs out in full ("Cluster randomised controlled trial",
# "Randomised controlled trial") rather than the abbreviated canonical labels
# ("Cluster RCT", "RCT"), so prefix matching against the dictionary silently
# stopped finding a match. Keyword matching is resilient to that kind of
# wording drift (and to typos like "controled").
study_design_tag <- function(value) {
  if (is.null(value) || length(value) == 0 || is.na(value) || !nzchar(trimws(as.character(value)))) {
    return(tags$span("Not reported"))
  }
  # Trim first: a value that starts with a line break would otherwise have an
  # empty "first line" and always be classified as "Other".
  raw <- trimws(as.character(value))
  first_line <- trimws(strsplit(raw, "\n")[[1]][1])
  display_text <- gsub("[\r\n]+", " ", raw)
  
  fl <- tolower(first_line)
  
  best_term <- dplyr::case_when(
    str_detect(fl, "pilot") ~ "Pilot trial",
    str_detect(fl, "non-?random") ~ "Non-randomised controlled trial",
    str_detect(fl, "cluster") ~ "Cluster RCT",
    str_detect(fl, "randomi[sz]ed controlled trial|\\brct\\b") ~ "RCT",
    str_detect(fl, "difference-in-differences|diff-in-diff") ~ "Difference-in-differences",
    str_detect(fl, "interrupted time series") ~ "Interrupted time series",
    str_detect(fl, "propensity score") ~ "Propensity score matching/weighted comparison",
    str_detect(fl, "regression discontinuity") ~ "Regression discontinuity",
    str_detect(fl, "pre-post|pre/post") ~ "Pre-post (single group, no comparison)",
    str_detect(fl, "cohort") ~ "Longitudinal cohort",
    str_detect(fl, "cross-sectional") ~ "Cross-sectional comparison",
    str_detect(fl, "quasi-experimental|ecological") ~ "Quasi-experimental study",
    TRUE ~ "Other"
  )
  
  tags$span(
    class = "sd-tooltip",
    `data-tooltip` = study_design_definitions[[best_term]],
    display_text
  )
}

# =============================================================================
# 11. Server definition
# =============================================================================

# Heading styles shared by every pop-up (previously copied into each one)
modal_title <- function(text) {
  tags$div(
    style = paste0(
      "font-family:'GT Pressura Extended', 'General Sans', sans-serif; ",
      "font-size:1.65rem; line-height:1.2; font-weight:700; color:", sfh_charcoal, ";"
    ),
    text
  )
}

modal_heading <- function(text, margin_top = "0", margin_bottom = "14px", bold = FALSE) {
  tags$h2(
    style = paste0(
      "font-family:'GT Pressura Extended', 'General Sans', sans-serif; ",
      "font-size:1.65rem; line-height:1.2; ",
      if (bold) "font-weight:700; ",
      "margin-top:", margin_top, "; margin-bottom:", margin_bottom, "; color:", sfh_charcoal, ";"
    ),
    text
  )
}

# The small "New" tag in the first column of both tables. Style guide colours:
# Pink fill with Charcoal text (white on pink is too low contrast at this size).
new_badge_html <- paste0(
  "<span style=\"background:", sfh_pink, "; color:", sfh_charcoal,
  "; padding:3px 6px; border-radius:10px; font-size:11px;\">New</span>"
)

# -----------------------------------------------------------------------------
# Status-line helpers
# -----------------------------------------------------------------------------
# The sentences above the tables ("... currently showing [91] evaluations ...")
# put each number in a coloured pill. Text colour follows the style guide via
# badge_text_colour() (charcoal on orange, white on green/blue).
status_base_style <- "font-size:15px; line-height:1.4; font-family:\"GT Pressura Extended\", sans-serif; color:#313131;"

status_pill <- function(value, fill) {
  paste0(
    "<span style='display:inline-block; padding:2px 8px; border-radius:999px; ",
    "background:", fill, "; color:", badge_text_colour(fill), "; font-weight:800;'>",
    value, "</span>"
  )
}

status_note <- paste0(
  "<span style='font-family:\"General Sans\", sans-serif; font-size:13px; color:#5A5A5A;'>",
  "(after filters are applied).</span>"
)

# -----------------------------------------------------------------------------
# Filter definitions
# -----------------------------------------------------------------------------
# Each dropdown filter is described once here and then used everywhere: to
# narrow the data, to test whether any filter is active, and to build the
# step-by-step counts in the "no matches" pop-ups.
#   step   - label shown in the pop-up
#   input  - id of the dropdown
#   kind   - "country": match any country in a multi-country cell
#            "value":   the column holds one value per row
#            "list":    the column holds a list of values per row; keep a row
#                       if any of them is selected
#   col    - the data column (not used for "country")
#   trim   - trim spaces before comparing (list columns only)
filter_spec <- function(step, input_id, kind, col = NULL, trim = FALSE) {
  list(step = step, input = input_id, kind = kind, col = col, trim = trim)
}

# Studies included tab (the order is the order of the steps in the pop-up)
study_filter_specs <- list(
  filter_spec("Country",                  "country_filter",          "country"),
  filter_spec("Region",                   "region_filter",           "list",  "World.Region_List"),
  filter_spec("Income classification",    "income_filter",           "value", "Income.Classification"),
  filter_spec("Outcome domain",           "outcome_domain_filter",   "list",  "Outcome_Domain_List"),
  filter_spec("Specific outcome",         "specific_outcome_filter", "list",  "Specific_Outcome_List"),
  filter_spec("Evidence rating",          "evidence_rating_filter",  "value", "evidence_rating"),
  filter_spec("Age Group",                "population_filter",       "list",  "Age.Groups_List"),
  filter_spec("Location of intervention", "context_by_filter",       "list",  "Level_List", trim = TRUE),
  filter_spec("INSPIRE strategies",       "inspire_filter",          "list",  "Inspire_List")
)

# Prevention explorer tab
int_filter_specs <- list(
  filter_spec("Region",                   "int_region_filter",  "list", "World.Region_List"),
  filter_spec("Country",                  "int_country_filter", "country"),
  filter_spec("Type of CSV outcome(s)",   "int_outcome_filter", "list", "Outcome_Domain_List", trim = TRUE),
  filter_spec("Age groups",               "int_age_filter",     "list", "Age.Groups_List",     trim = TRUE),
  filter_spec("Contexts where the intervention was implemented",
                                          "int_level_filter",   "list", "Level_List",          trim = TRUE),
  filter_spec("INSPIRE category",         "int_inspire_filter", "list", "Inspire_List")
)

# Narrow `d` to the rows that match the values `selected` in one filter
# (no selection = no change).
apply_filter_spec <- function(d, spec, selected) {
  if (length(selected) == 0) return(d)
  keep <- switch(
    spec$kind,
    country = country_matches_selection(d$Country, selected),
    value   = d[[spec$col]] %in% selected,
    list    = lengths(lapply(d[[spec$col]], function(x) {
      if (spec$trim) intersect(trimws(x), trimws(selected)) else intersect(x, selected)
    })) > 0
  )
  d[keep, , drop = FALSE]
}

server <- function(input, output, session) {
  
  session_token <- session$token
  session_start <- Sys.time()
  log_event("page_view", paste0("session_", session_token))
  session$onSessionEnded(function() {
    secs <- round(as.numeric(difftime(Sys.time(), session_start, units = "secs")))
    log_event("session_end", paste0("session_", session_token, "; seconds=", secs))
  })
  session$sendCustomMessage("init-dt-view-button", list())
  
  # Log which filters are used and which values are picked. Every filter is a
  # fixed-choice dropdown, so only known category names are ever recorded (never
  # free text). Clearing a filter is not logged.
  filter_ids <- c(
    "country_filter", "region_filter", "income_filter",
    "outcome_domain_filter", "specific_outcome_filter",
    "evidence_rating_filter", "population_filter",
    "context_by_filter", "inspire_filter",
    "int_region_filter", "int_country_filter", "int_outcome_filter",
    "int_age_filter", "int_level_filter", "int_inspire_filter"
  )
  for (filter_id in filter_ids) {
    local({
      fid <- filter_id
      observeEvent(input[[fid]], {
        log_event("filter", paste0(fid, ": ", paste(input[[fid]], collapse = "; ")))
      }, ignoreInit = TRUE, ignoreNULL = TRUE)
    })
  }
  
  citation_text_dashboard <- paste0(
    "Department of Social Policy and Intervention and Safe Futures Hub. 2026. ",
    "An interactive dashboard on interventions to prevent child sexual violence. ",
    "Available from: https://safefutureshub.shinyapps.io/CSVlivingsystematicreview/"
  )
  
  # Apply every filter in `specs` to `d`, using the current dropdown values.
  filtered_by <- function(d, specs) {
    for (spec in specs) d <- apply_filter_spec(d, spec, input[[spec$input]])
    d
  }

  any_filter_active <- function(specs) {
    any(vapply(specs, function(spec) length(input[[spec$input]]) > 0, logical(1)))
  }

  # The numbers behind a "no matches" pop-up: the count before any filter, then
  # the count after each filter that is switched on, in order. `count_fn` says
  # how to count (rows, or distinct interventions).
  filter_steps_table <- function(start, specs, count_fn, first_label) {
    out <- list(data.frame(step = first_label, n = count_fn(start), stringsAsFactors = FALSE))
    current <- start
    for (spec in specs) {
      selected <- input[[spec$input]]
      if (length(selected) > 0) {
        current <- apply_filter_spec(current, spec, selected)
        out[[length(out) + 1]] <- data.frame(
          step = paste0(spec$step, " (", paste(selected, collapse = ", "), ")"),
          n = count_fn(current),
          stringsAsFactors = FALSE
        )
      }
    }
    bind_rows(out)
  }

  # ---------------------------------------------------------------------------
  # Intervention-level dataset (effective interventions only)
  # ---------------------------------------------------------------------------
  intervention_summary <- reactive({
    
    d <- filtered_data()
    
    # Pre-clean the materials-availability fields once, row-by-row, so the
    # summary below can pick availability + link from the *same* evaluation
    # row instead of picking each independently (which could pair one
    # evaluation's availability with a different evaluation's link).
    d$.avail_clean <- if ("Intervention.Materials.Availability" %in% names(d)) {
      trimws(as.character(d$Intervention.Materials.Availability))
    } else {
      NA_character_
    }
    
    d$.avail_ok <- !is.na(d$.avail_clean) &
      nzchar(d$.avail_clean) &
      !tolower(d$.avail_clean) %in% c("na", "n/a", "none")
    
    d$.link_clean <- if ("Link.to.materials" %in% names(d)) {
      trimws(as.character(d$Link.to.materials))
    } else {
      NA_character_
    }
    
    # Check whether this row contains a real website address.
    d$.link_ok <- !is.na(d$.link_clean) &
      nzchar(d$.link_clean) &
      !tolower(d$.link_clean) %in% c("na", "n/a", "none") &
      stringr::str_detect(d$.link_clean, "https?://")
    
    # Give the best rows the smallest number:
    # 1 = a real link
    # 2 = an availability description but no link
    # 3 = no useful information
    d$.materials_rank <- dplyr::case_when(
      d$.link_ok  ~ 1L,
      d$.avail_ok ~ 2L,
      TRUE        ~ 3L
    )
    
    # --- Intervention-page specific filters ----------------------------------
    d <- filtered_by(d, int_filter_specs)

    # keep only effective studies for prevention explorer
    d <- d %>%
      filter(intervention_effectiveness_code == "R")
    
    
    
    # --- Summarise to intervention level -------------------------------------
    # Intervention.ID is the canonical intervention key (e.g. the three
    # "None (Honeypot website and messaging)" rows carry IDs 35/36/37
    # because they are three different interventions that happen to share a
    # placeholder name) - group by ID+Name so each distinct ID is its own row.
    d %>%
      group_by(Intervention.ID, Intervention.Name) %>%
      summarise(
        Intervention.ID = {
          x <- unique(na.omit(as.character(Intervention.ID)))
          if (length(x) == 0) NA_character_ else x[1]
        },

        Evaluations = n(),
        
        # Mark this intervention as New when at least one
        # of its evaluation rows has a new update date.
        Is.New = any(is_new_update, na.rm = TRUE),
        # Split multi-country cells (e.g. "Kenya, South Africa") before taking
        # the unique set, so a country already covered by one evaluation row
        # doesn't also show up as part of a combined-cell string from another
        # (e.g. "Kenya, South Africa, Kenya").
        Countries = paste(unique(unlist(lapply(na.omit(Country), split_country_tokens))), collapse = ", "),
        Regions = paste(unique(na.omit(World.Region)), collapse = ", "),
        # --- Plain-text fields (feed the CSV export directly, and the HTML
        # badge columns below, so both stay in sync with a single source). ---
        Target = {
          x <- if ("Intervention.Target" %in% names(.)) pick_richest_value(Intervention.Target) else NA_character_
          ifelse(!is.na(x), x, "Not reported")
        },
        Study.Design = {
          x <- if ("Study.Design" %in% names(.)) first(Study.Design) else NA_character_
          ifelse(!is.na(x) && nzchar(x), x, "Not reported")
        },
        Outcomes_plain = make_tag_plain(
          unique(unlist(Outcome_Domain_List)),
          order = c(
            "Victimisation",
            "Perpetration",
            "Both (combined vic/perp outcome)"
          )
        ),
        Age.Groups_plain = make_tag_plain(
          unique(unlist(Age.Groups_List)),
          order = age_group_order
        ),
        Inspire_plain = make_tag_plain(
          unique(unlist(Inspire_List)),
          order = inspire_category_order
        ),
        Context_plain = {
          vals <- unique(unlist(Level_List))
          vals <- vals[!is.na(vals) & nzchar(vals)]

          vals <- c(intersect(context_order, vals), setdiff(vals, context_order))

          if (length(vals) == 0) NA_character_ else paste(vals, collapse = ", ")
        },

        # --- HTML badge fields (used by the interactive table only) ---
        Outcomes = make_tag_badges_html(
          unique(unlist(Outcome_Domain_List)),
          order = c(
            "Victimisation",
            "Perpetration",
            "Both (combined vic/perp outcome)"
          ),
          special_fill = c(
            "Both (combined vic/perp outcome)" = sfh_pink
          )
        ),
        Age.Groups = make_tag_badges_html(
          unique(unlist(Age.Groups_List)),
          order = age_group_order,
          special_fill = c(
            "NA" = sfh_sand
          ),
          special_text = c(
            "NA" = sfh_charcoal
          )
        ),
        Inspire = make_tag_badges_html(
          unique(unlist(Inspire_List)),
          order = inspire_category_order
        ),
        Context = {

          vals <- unique(unlist(Level_List))
          vals <- vals[!is.na(vals) & nzchar(vals)]

          if (length(vals) == 0) {
            NA_character_
          } else {

            vals <- c(intersect(context_order, vals), setdiff(vals, context_order))

            badge_palette <- c(sfh_green, sfh_blue, sfh_purple, sfh_orange, sfh_ocean, sfh_pink)

            # Colour keyed to position in context_order (not this row's local
            # index) so a given location always renders the same colour.
            make_badge <- function(label) {
              col <- badge_colour_for(label, context_order, badge_palette)
              make_uniform_badge(label, col)
            }

            badges <- vapply(vals, make_badge, FUN.VALUE = character(1))

            paste0(
              "<div style='display:flex; flex-direction:column; gap:3px;'>",
              paste(badges, collapse = ""),
              "</div>"
            )
          }
        },
        Availability = {
          idx <- which.min(.materials_rank)
          # which.min() returns integer(0) when the group has no rows, which
          # happens during dplyr's zero-row type-inference pass whenever a
          # filter selection (e.g. a World Region with no "R"-rated
          # interventions) leaves nothing to summarise. Treat that the same
          # as "no usable value" instead of indexing into an empty vector.
          v <- if (length(idx) == 0) NA_character_ else .avail_clean[idx]

          if (
            is.na(v) ||
            !nzchar(v) ||
            tolower(v) %in% c("na", "n/a", "none")
          ) {
            NA_character_
          } else {
            v
          }
        },

        MaterialsLink = {
          idx <- which.min(.materials_rank)
          v <- if (length(idx) == 0) NA_character_ else .link_clean[idx]

          if (
            is.na(v) ||
            !nzchar(v) ||
            tolower(v) %in% c("na", "n/a", "none") ||
            !stringr::str_detect(v, "https?://")
          ) {
            NA_character_
          } else {
            v
          }
        },
        .groups = "drop"
      ) %>%
      arrange(desc(Evaluations))
  })
  
  summary_counts_interventions_fixed <- reactive({
    d <- df
    
    d <- d %>%
      filter(intervention_effectiveness_code == "R")
    
    intervention_level <- d %>%
      # Intervention.ID is the canonical intervention key - see the matching
      # note in intervention_summary() above.
      group_by(Intervention.ID, Intervention.Name) %>%
      summarise(
        Evaluations = n_distinct(Study.ID),
        # Split multi-country cells (e.g. "Kenya, South Africa") before taking
        # the unique set, so a country already covered by one evaluation row
        # doesn't also show up as part of a combined-cell string from another
        # (e.g. "Kenya, South Africa, Kenya").
        Countries = paste(unique(unlist(lapply(na.omit(Country), split_country_tokens))), collapse = ", "),
        .groups = "drop"
      )
    
    list(
      n_evals = sum(intervention_level$Evaluations, na.rm = TRUE),
      n_interventions = nrow(intervention_level),
      n_countries = count_countries(intervention_level$Countries)
    )
  })
  
  intervention_summary_counts <- reactive({
    d <- intervention_summary()
    
    n_interventions <- nrow(d)
    n_evals <- sum(d$Evaluations, na.rm = TRUE)
    n_countries <- count_countries(d$Countries)
    
    list(
      n_interventions = n_interventions,
      n_evals = n_evals,
      n_countries = n_countries
    )
  })
  
  summary_counts_static <- reactive({
    d <- df
    
    n_evals <- if ("Study.ID" %in% colnames(d)) {
      d %>%
        filter(!is.na(Study.ID)) %>%
        distinct(Study.ID) %>%
        nrow()
    } else {
      nrow(d)
    }
    
    n_interventions <- if ("Intervention.Name" %in% colnames(d)) {
      d %>%
        filter(!is.na(Intervention.Name)) %>%
        distinct(Intervention.ID, Intervention.Name) %>%
        nrow()
    } else {
      NA_integer_
    }

    # Split multi-country cells before counting distinct countries (a cell like
    # "Kenya, South Africa" is 2 countries, not 1), matching the approach used
    # for the Prevention explorer's country count (see
    # summary_counts_interventions_fixed()) so the two tallies stay comparable.
    n_countries <- if ("Country" %in% colnames(d)) {
      count_countries(d$Country)
    } else {
      NA_integer_
    }

    list(
      n_evals = n_evals,
      n_interventions = n_interventions,
      n_countries = n_countries
    )
  })
  
  output$active_filter_chips <- renderUI({
    
    make_chip <- function(label, input_id) {
      tags$span(
        class = "filter-chip",
        `data-value` = label,
        `data-input` = input_id,
        style = "
        background:#F3F2E8;
        border:1px solid #C2C2BA;
        padding:4px 10px;
        border-radius:999px;
        font-size:12px;
        font-family:'General Sans';
        display:inline-flex;
        align-items:center;
        gap:6px;
        cursor:pointer;
      ",
        label,
        tags$span("\u00d7", style = "font-weight:bold;")
      )
    }
    
    filter_map <- list(
      "country_filter" = input$country_filter,
      "region_filter" = input$region_filter,
      "income_filter" = input$income_filter,
      "outcome_domain_filter" = input$outcome_domain_filter,
      "specific_outcome_filter" = input$specific_outcome_filter,
      "evidence_rating_filter" = input$evidence_rating_filter,
      "population_filter" = input$population_filter,
      "context_by_filter" = input$context_by_filter,
      "inspire_filter" = input$inspire_filter
    )
    
    chips <- list()
    
    for (input_id in names(filter_map)) {
      vals <- filter_map[[input_id]]
      
      if (!is.null(vals) && length(vals) > 0) {
        chips <- c(chips, lapply(vals, function(x) {
          make_chip(x, input_id)
        }))
      }
    }
    
    if (length(chips) == 0) return(NULL)
    
    tags$div(
      style = "margin:10px 0 6px 0;",
      tags$div(
        style = "font-size:12px; color:#5A5A5A; margin-bottom:4px;",
        "Active filters:"
      ),
      tags$div(
        style = "display:flex; flex-wrap:wrap; gap:6px;",
        chips
      )
    )
  })
  
  output$status_line_evidence_static <- renderUI({
    sc <- summary_counts_static()
    
    HTML(paste0(
      "<div style='", status_base_style, " margin:8px 0 14px 0; padding:12px 14px; background:#F3F2E8; border-left:5px solid #F09E56; border-radius:12px;'>",
      "This page contains a total of ", status_pill(sc$n_evals, sfh_orange),
      " evaluations conducted among ", status_pill(sc$n_interventions, sfh_green),
      " interventions in ", status_pill(sc$n_countries, sfh_blue),
      " countries.",
      "</div>"
    ))
  })
  
  int_has_active_filters <- reactive({
    any_filter_active(int_filter_specs)
  })

  int_filter_step_counts <- reactive({
    d0 <- filtered_data()

    # Intervention tab only shows effective studies
    current <- d0 %>% filter(intervention_effectiveness_code == "R")

    # current is evaluation-level (one row per evaluation); count distinct
    # interventions at each step so these numbers line up with the
    # intervention counts shown above/below the table (intervention_summary_counts()).
    # Intervention.ID is the canonical intervention key (some interventions
    # share a placeholder Intervention.Name but are different interventions
    # with different IDs), so distinct on ID+Name to match intervention_summary().
    n_distinct_interventions <- function(d) {
      d %>% filter(!is.na(Intervention.Name)) %>% distinct(Intervention.ID, Intervention.Name) %>% nrow()
    }

    filter_steps_table(current, int_filter_specs, n_distinct_interventions,
                       "All effective interventions")
  })

  # Show a "No matched ..." pop-up, with the step-by-step counts, when the active
  # filters leave nothing; dismiss it again when the filters change or are cleared.
  #   active - reactive: is any filter switched on?
  #   data   - reactive: the rows the tab is showing
  #   steps  - reactive: the step-by-step counts table
  #   noun   - "interventions" or "studies"
  #   hint   - the advice sentence in the pop-up
  watch_no_matches <- function(active, data, steps, noun, title, hint) {
    shown <- reactiveVal(FALSE)

    observe({
      if (!active()) {
        # Dismiss a modal left open from an earlier filter combination so it
        # doesn't keep showing stale step counts once filters are cleared.
        if (isTRUE(shown())) removeModal()
        shown(FALSE)
        return()
      }

      # Guarded: unlike renderUI/renderDT/downloadHandler (which Shiny catches
      # per-output), an error inside a bare observe() propagates and kills the
      # whole session ("disconnected from server") instead of just failing this
      # pop-up, so catch and log it.
      d <- tryCatch(data(), error = function(e) {
        message("no_match_popup (", noun, "): data failed: ", conditionMessage(e))
        NULL
      })
      if (is.null(d)) return()

      if (nrow(d) > 0) {
        # A filter change that brings back matches should close a previously
        # shown pop-up instead of leaving stale step counts on screen.
        if (isTRUE(shown())) removeModal()
        shown(FALSE)
        return()
      }

      if (isTRUE(shown())) return()

      step_df <- tryCatch(steps(), error = function(e) {
        message("no_match_popup (", noun, "): step counts failed: ", conditionMessage(e))
        NULL
      })
      if (is.null(step_df)) return()

      step_prefix <- ifelse(
        seq_along(step_df$step) == 1,
        "Without filters: ",
        paste0("After filter ", seq_along(step_df$step) - 1, " was applied: ")
      )

      step_html <- paste0(
        "<ul style='padding-left:18px; list-style:none;'>",
        paste0(
          "<li><strong>",
          htmltools::htmlEscape(paste0(step_prefix, step_df$step)),
          "</strong> &rarr; ",
          step_df$n,
          " ", noun, "</li>",
          collapse = ""
        ),
        "</ul>"
      )

      showModal(
        modalDialog(
          title = title,
          easyClose = TRUE,
          footer = modalButton("Close"),
          HTML(
            paste0(
              "<p>No ", noun, " match the current combination of filters.</p>",
              "<p>", hint, "</p>",
              "<p><strong>How the number of ", noun, " changed based on the filters you applied:</strong></p>",
              step_html
            )
          )
        )
      )

      shown(TRUE)
    })
  }

  watch_no_matches(
    int_has_active_filters, intervention_summary, int_filter_step_counts,
    noun = "interventions", title = "No matched interventions",
    hint = "Try removing one or two filters to broaden the results, or click <strong>Reset </strong> to start again."
  )

  
  output$status_line_intervention_static <- renderUI({
    sc <- summary_counts_interventions_fixed()
    
    HTML(paste0(
      "<div style='", status_base_style, " margin-top:8px; margin-bottom:12px;'>",
      "Prevention efforts that are highlighted in this tab have at least one evaluation since 2018 that found a reduction in at least one CSV victimisation or perpetration outcome and no increase in any CSV victimisation and perpetration outcomes.<br>",
      "This prevention explorer contains a total of ", status_pill(sc$n_interventions, sfh_orange),
      " effective interventions, evaluated by ", status_pill(sc$n_evals, sfh_green),
      " evaluations in ", status_pill(sc$n_countries, sfh_blue),
      " countries.",
      "</div>"
    ))
  })
  
  output$status_line_intervention_explorer <- renderUI({
    sc <- intervention_summary_counts()
    
    HTML(paste0(
      "<div style='", status_base_style, " margin-top:8px; margin-bottom:8px;'>",
      "The below interactive table is currently showing ", status_pill(sc$n_interventions, sfh_orange),
      " effective interventions, evaluated by ", status_pill(sc$n_evals, sfh_green),
      " evaluations in ", status_pill(sc$n_countries, sfh_blue),
      " countries ", status_note,
      "</div>"
    ))
  })
  
  has_active_filters <- reactive({
    any_filter_active(study_filter_specs)
  })

  filter_step_counts <- reactive({
    filter_steps_table(df, study_filter_specs, nrow, "All studies")
  })

  # ---------------------------------------------------------------------------
  # Reactive filtered data (used by both tabs)
  # ---------------------------------------------------------------------------
  filtered_data <- reactive({
    filtered_by(df, study_filter_specs)
  })
  
  watch_no_matches(
    has_active_filters, filtered_data, filter_step_counts,
    noun = "studies", title = "No matched studies",
    hint = "You may want to remove one or two filters to broaden the results, or click <strong>Reset </strong> to start again."
  )
  
  # ---------------------------------------------------------------------------
  # Download displayed columns (Studies included)
  # ---------------------------------------------------------------------------
  output$download_all <- downloadHandler(
    filename = function() {
      paste0("CSV_vicperp_study_summary_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      d <- normalise_df(filtered_data())
      d <- d[, intersect(csv_export_columns, names(d)), drop = FALSE]
      d <- clean_export_df(d)
      write_clean_xlsx(d, file, "studies")
    }
  )
  
  # ---------------------------------------------------------------------------
  # Global status-line counts (filtered)
  # ---------------------------------------------------------------------------
  summary_counts <- reactive({
    d <- filtered_data()
    
    if ("Study.ID" %in% colnames(d)) {
      n_evals <- d %>%
        filter(!is.na(Study.ID)) %>%
        distinct(Study.ID) %>%
        nrow()
    } else {
      n_evals <- nrow(d)
    }
    
    if ("Intervention.Name" %in% colnames(d)) {
      n_interventions <- d %>%
        filter(!is.na(Intervention.Name)) %>%
        distinct(Intervention.ID, Intervention.Name) %>%
        nrow()
    } else {
      n_interventions <- NA_integer_
    }

    # Split multi-country cells before counting distinct countries (see
    # summary_counts_static() for the same fix and why it's needed).
    n_countries <- count_countries(d$Country)

    list(
      n_evals         = n_evals,
      n_interventions = n_interventions,
      n_countries     = n_countries
    )
  })
  
  
  output$status_line_explorer <- renderUI({
    sc <- summary_counts()
    
    HTML(paste0(
      "<div style='", status_base_style, " margin-top:8px;'>",
      "The below interactive table is currently showing ", status_pill(sc$n_evals, sfh_orange),
      " evaluations conducted among ", status_pill(sc$n_interventions, sfh_green),
      " interventions, in ", status_pill(sc$n_countries, sfh_blue),
      " countries ", status_note,
      "</div>"
    ))
  })
  
  
  # ---------------------------------------------------------------------------
  # Prevention explorer table
  # ---------------------------------------------------------------------------
  output$intervention_table <- renderDT({
    d <- intervention_summary()
    req(d, nrow(d) > 0)
    
    
    clean_pipe <- function(x) {
      if (is.null(x)) return("")
      x <- as.character(x)
      x <- gsub("\\|+", ", ", x)
      x <- gsub(",\\s*,", ", ", x)
      x <- gsub("^,\\s*|,\\s*$", "", x)
      trimws(x)
    }
    
    # Some "Link to materials" cells wrap the URL in prose, e.g.
    # "1. Curriculum (https://example.org/page)" - a bare [^\s]+ match would
    # swallow the trailing ")" into the URL and break the link. Strip
    # trailing punctuation that isn't actually part of the URL, but leave a
    # closing ")" alone if the URL itself contains a matching "(".
    clean_extracted_url <- function(u) {
      # perl = TRUE: R's default (POSIX ERE) engine does not treat "\]"
      # inside a character class as an escaped literal "]" the way PCRE
      # does, so this needs the PCRE engine to match as intended.
      u <- sub("[.,;:'\"\\]]+$", "", u, perl = TRUE)
      opens  <- lengths(regmatches(u, gregexpr("\\(", u)))
      closes <- lengths(regmatches(u, gregexpr("\\)", u)))
      while (nzchar(u) && closes > opens && grepl("\\)$", u)) {
        u <- sub("\\)$", "", u)
        closes <- closes - 1
      }
      u
    }
    
    # Availability badge, with a "materials" link appended when the source
    # "Link to materials" cell holds an actual URL (not blank / N/A / "None").
    make_availability_cell <- function(avail, link) {
      avail_label <- if (is.na(avail) || !nzchar(trimws(as.character(avail)))) {
        "Not described"
      } else {
        trimws(as.character(avail))
      }
      
      url <- NA_character_
      if (!is.na(link) && nzchar(trimws(as.character(link)))) {
        m <- regmatches(
          as.character(link),
          regexpr("https?://[^[:space:]]+", as.character(link), perl = TRUE)
        )
        if (length(m) > 0 && nzchar(m[1])) url <- clean_extracted_url(m[1])
      }
      
      fill <- dplyr::case_when(
        grepl("^yes", tolower(avail_label))  ~ sfh_green,
        grepl("^no",  tolower(avail_label))  ~ sfh_orange,
        TRUE                                 ~ sfh_sand
      )
      textc <- badge_text_colour(fill)
      
      pill_style <- paste0(
        "display:inline-block; padding:3px 8px; border-radius:999px; ",
        "background:", fill, "; color:", textc, "; font-family:\"General Sans\"; font-size:11px;"
      )
      
      if (is.na(url)) {
        paste0("<span style='", pill_style, "'>", htmltools::htmlEscape(avail_label), "</span>")
      } else {
        # The whole pill becomes the clickable link (no separate "link" text) -
        # a trailing arrow signals it's clickable.
        paste0(
          "<a href='", htmltools::htmlEscape(url, attribute = TRUE),
          "' target='_blank' rel='noopener noreferrer' title='Open materials link' ",
          "style='", pill_style, " text-decoration:none; cursor:pointer;'>",
          htmltools::htmlEscape(avail_label), " &#8599;",
          "</a>"
        )
      }
    }
    
    final_df <- d %>%
      mutate(
        New = ifelse(Is.New, new_badge_html, ""),
        .view_id = ifelse(
          !is.na(Intervention.ID) & nzchar(trimws(as.character(Intervention.ID))),
          paste0("intid::", as.character(Intervention.ID)),
          paste0("intname::", URLencode(as.character(Intervention.Name), reserved = TRUE))
        ),
        View = sprintf(
          '<button class="btn btn-sm btn-primary view-btn" data-id="%s">View</button>',
          htmltools::htmlEscape(.view_id, attribute = TRUE)
        ),
        Countries = sapply(Countries, clean_pipe),
        Context = sapply(Context, clean_pipe),
        Availability = mapply(make_availability_cell, Availability, MaterialsLink, USE.NAMES = FALSE)
      ) %>%
      select(
        New,
        View,
        `Intervention name` = Intervention.Name,
        Region = Regions,
        Country = Countries,
        `Type of CSV outcome(s)` = Outcomes,
        `Included age groups` = Age.Groups,
        `Location of<br/>implementation` = Context,
        `INSPIRE category` = Inspire,
        `# Evaluations since 2018` = Evaluations,
        `Public access<br/>to materials` = Availability
      )
    
    datatable(
      final_df,
      escape = FALSE,
      rownames = FALSE,
      selection = "none",
      options = list(
        scrollX = TRUE,
        scrollY = "55vh",
        scrollCollapse = TRUE,
        pageLength = 25,
        autoWidth = FALSE,
        dom = "frtip",
        
        # Allow users to sort the useful columns.
        ordering = TRUE,
        
        # Starting order:
        # 1. Number of evaluations, highest to lowest
        # 2. Intervention name, A-Z when evaluation counts are equal
        #
        # DataTables starts counting columns at zero:
        # 9 = Number of evaluations
        # 2 = Intervention name
        order = list(
          list(9, "desc"),
          list(2, "asc")
        ),

        columnDefs = list(
          list(width = "50px",  targets = 0),   # New
          list(width = "70px",  targets = 1),   # View
          list(width = "190px", targets = 2),   # Intervention name
          list(width = "130px", targets = 3),   # Region
          list(width = "140px", targets = 4),   # Country
          list(width = "190px", targets = 5),   # Type of CSV outcome(s)
          list(width = "190px", targets = 6),   # Included age groups
          list(width = "150px", targets = 7),   # Location of implementation
          list(width = "170px", targets = 8),   # INSPIRE category
          list(width = "110px", targets = 9),   # # Evaluations since 2018
          list(width = "170px", targets = 10),  # Public access

          # Do not show sorting arrows for New and View.
          list(
            orderable = FALSE,
            targets = c(0, 1)
          ),
          
          # Centre-align every column, both header and cell text.
          list(
            className = "dt-center",
            targets = "_all"
          )
        )
      )
    )
  })
  
  output$download_intervention_summary <- downloadHandler(
    filename = function() {
      paste0("CSV_vicperp_intervention_summary_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      # Use the plain-text sibling columns (not the HTML badge columns used
      # by the interactive table) so the export shows readable values like
      # "Victimisation, Perpetration" instead of raw <span> markup.
      export_df <- intervention_summary() %>%
        select(
          `Intervention name` = Intervention.Name,
          Region = Regions,
          Country = Countries,
          `Type of CSV outcome(s)` = Outcomes_plain,
          `Included age groups` = Age.Groups_plain,
          `INSPIRE category` = Inspire_plain,
          `Location of implementation` = Context_plain,
          `Target population` = Target,
          `Study design` = Study.Design,
          `# Evaluations since 2018` = Evaluations,
          `Public access to materials` = Availability,
          `Materials link` = MaterialsLink
        )
      write_clean_xlsx(export_df, file, "interventions")
    }
  )
  
  
  # ---------------------------------------------------------------------------
  # Home / Studies included static plots (operate on the full df)
  # ---------------------------------------------------------------------------
  make_counts_home_gg <- reactive({
    d <- normalise_df(df)
    
    if ("Study.ID" %in% names(d)) {
      n_evals <- d %>% filter(!is.na(Study.ID)) %>% distinct(Study.ID) %>% nrow()
    } else {
      n_evals <- nrow(d)
    }
    if ("Intervention.Name" %in% names(d)) {
      n_interventions <- d %>% filter(!is.na(Intervention.Name)) %>% distinct(Intervention.ID, Intervention.Name) %>% nrow()
    } else {
      n_interventions <- NA_integer_
    }
    n_countries <- count_countries(d$Country)

    # Build counts table in visual order: countries, interventions, evaluations
    df_counts <- tibble(
      Metric = c("Countries represented", "Interventions included", "Unique evaluations"),
      Value  = c(n_countries, n_interventions, n_evals)
    ) %>% mutate(Value = replace_na(Value, 0L))
    
    df_counts$Metric <- factor(df_counts$Metric, levels = rev(c(
      "Countries represented", "Interventions included", "Unique evaluations"
    )))
    
    max_val <- max(df_counts$Value, 1)
    if (max_val <= 20) step <- 1 else if (max_val <= 50) step <- 5 else step <- ceiling(max_val / 10)
    
    ggplot(df_counts, aes(x = Value, y = Metric)) +
      geom_col(width = 0.6, fill = dark_purple) +
      scale_x_continuous(breaks = seq(0, max_val, by = step), limits = c(0, max_val + step)) +
      labs(title = " ", x = "Count", y = "") +
      theme_sfh()
  })
  
  make_age_groups_gg <- reactive({
    # Use the data as it is, NOT normalise_df(): that would flatten each study's
    # list of age groups into one text string ("Adolescence; Adults"), which the
    # classification below would then count under only one group.
    d <- df
    
    if (!"Age.Groups_List" %in% names(d)) {
      d$Age.Groups_List <- parse_multi_values(d$Age.Group.Description)
    }
    
    # Expand so one study can contribute to multiple age categories
    age_long <- d %>%
      mutate(rowid = row_number()) %>%
      select(rowid, Study.ID, Age.Groups_List) %>%
      unnest_longer(Age.Groups_List, values_to = "age_raw") %>%
      mutate(age_group = classify_age_group(age_raw)) %>%
      filter(!is.na(age_group)) %>%
      distinct(rowid, age_group) %>%
      count(age_group, name = "n") %>%
      mutate(
        age_group = factor(
          age_group,
          levels = rev(age_group_order)
        )
      ) %>%
      arrange(age_group)
    
    if (nrow(age_long) == 0) {
      ggplot() + theme_void() + labs(caption = "No age group data")
    } else {
      ggplot(age_long, aes(x = age_group, y = n)) +
        geom_col(fill = orange_main, width = 0.7) +
        coord_flip() +
        labs(x = "", y = "Count") +
        theme_sfh()
    }
  })
  
  make_income_home_gg <- reactive({
    d <- normalise_df(df)
    if (!"Country" %in% names(d))               d$Country <- NA_character_
    if (!"Income.Classification" %in% names(d)) d$Income.Classification <- NA_character_
    
    d <- d %>%
      mutate(Country = as.character(Country),
             Income.Classification = as.character(Income.Classification))
    
    tab <- country_income_table(d) %>%
      filter(!is.na(income_std)) %>%
      count(income_std, name = "n")
    
    tab <- tibble(income_std = income_levels) %>%
      left_join(tab, by = "income_std") %>%
      mutate(n = replace_na(n, 0))
    
    maxn <- max(tab$n, 1)
    
    ggplot(tab, aes(x = reorder(income_std, n), y = n)) +
      geom_col(width = 0.7, fill = teal_main) +
      coord_flip() +
      scale_y_continuous(breaks = seq(0, maxn, 1), limits = c(0, maxn + 0.5)) +
      labs(x = "", y = "Number of countries") +
      theme_sfh()
  })
  
  make_timeline_home_gg <- reactive({
    d <- normalise_df(df)
    if (!"publication_year" %in% names(d)) d$publication_year <- NA_integer_
    d$publication_year <- suppressWarnings(as.integer(as.character(d$publication_year)))
    if (!"Outcome_Domain_List" %in% names(d)) d$Outcome_Domain_List <- NA_character_

    d <- d %>% filter(!is.na(publication_year))

    if (nrow(d) == 0) {
      return(ggplot() + theme_void() + labs(caption = "No publication year data"))
    }

    domain_text <- ifelse(is.na(d$Outcome_Domain_List), "", d$Outcome_Domain_List)
    is_perp <- grepl("Perpetration",  domain_text, fixed = TRUE)
    is_vic  <- grepl("Victimisation", domain_text, fixed = TRUE)

    year_seq <- seq(min_year, max_year, by = 1)
    count_by_year <- function(flag) {
      as.integer(table(factor(d$publication_year[flag], levels = year_seq)))
    }

    timeline_series_colors <- c(
      "All prevention efforts" = sfh_charcoal,
      "CSV perpetration"       = sfh_orange,
      "CSV victimisation"      = sfh_blue
    )

    d3 <- tibble(
      year   = rep(year_seq, 3),
      series = rep(names(timeline_series_colors), each = length(year_seq)),
      n      = c(count_by_year(rep(TRUE, nrow(d))), count_by_year(is_perp), count_by_year(is_vic))
    ) %>%
      mutate(series = factor(series, levels = names(timeline_series_colors))) %>%
      arrange(series, year) %>%
      group_by(series) %>%
      mutate(cum_n = cumsum(n)) %>%
      ungroup()

    y_max    <- max(10, ceiling(max(d3$cum_n, na.rm = TRUE) / 10) * 10)
    y_breaks <- seq(0, y_max, by = 10)

    end_labels <- d3 %>% filter(year == max_year)

    ggplot(d3, aes(x = year, y = cum_n, colour = series, group = series)) +
      geom_line(linewidth = 1.1) +
      geom_point(size = 2.4) +
      geom_text(
        data = end_labels,
        aes(x = max_year + 0.25, y = cum_n, label = series, colour = series),
        hjust = 0, size = 3.2, family = "General Sans", fontface = "plain",
        show.legend = FALSE
      ) +
      scale_colour_manual(values = timeline_series_colors) +
      scale_x_continuous(
        breaks = year_seq,
        limits = c(min_year - 0.4, max_year + 2.6)
      ) +
      scale_y_continuous(breaks = y_breaks, limits = c(0, y_max)) +
      coord_cartesian(clip = "off") +
      labs(x = "Year of evaluation", y = NULL) +
      theme_sfh(base_size = 13) +
      theme(
        legend.position     = "none",
        panel.grid.major.x  = element_blank(),
        panel.grid.major.y  = element_line(colour = sfh_sand_dark, linewidth = 0.4),
        axis.title.x        = element_text(size = 10, colour = sfh_charcoal_mid),
        axis.text           = element_text(size = 9),
        plot.margin         = margin(t = 8, r = 78, b = 8, l = 8)
      )
  })
  
  # ---------------------------------------------------------------------------
  # Plot outputs (full df)
  # ---------------------------------------------------------------------------
  output$timeline_plotly <- renderPlot({ make_timeline_home_gg() })
  output$age_plotly      <- renderPlot({ make_age_groups_gg() })
  output$income_plotly   <- renderPlot({ make_income_home_gg() })
  output$counts_bar      <- renderPlot({ make_counts_home_gg() })
  
  # Pretty effectiveness label (used in study modal)
  
  
  # ---------------------------------------------------------------------------
  # Country summary + leaflet map
  # ---------------------------------------------------------------------------
  # The world map on the Studies included tab shows the full dataset; it is not
  # linked to the sidebar filters.
  country_summary_home_reactive <- reactive({
    country_summary_for_map(df)
  })
  
  output$world_map_home <- renderLeaflet({
    render_country_map(country_summary_home_reactive(),
                       legend_title = "Number of evaluations",
                       zoom = 2.5)
  })
  
  # Pretty-print a country name for the modal title
  pretty_country <- function(x) {
    x <- as.character(x)
    x <- gsub("[\r\n\t]+", " ", x)
    x <- trimws(x)
    x <- gsub("\\s+", " ", x)
    
    x_low <- tolower(x)
    
    if (x_low %in% c("usa", "u.s.a.", "united states", "united states of america")) {
      return("United States of America")
    }
    if (x_low %in% c("uk", "u.k.", "united kingdom", "great britain")) {
      return("United Kingdom")
    }
    if (x_low %in% c("drc", "d.r.c.", "democratic republic of congo", "dr congo", "congo drc")) {
      return("Congo, Dem. Rep.")
    }
    
    tools::toTitleCase(x)
  }
  
  # Map-click handler: opens the country pop-up for the clicked country
  map_click_handler <- function(country_clicked) {
    if (is.null(country_clicked) || !nzchar(country_clicked)) return()
    
    # The map shows the whole dataset, so this looks at every study.
    country_df <- normalise_df(df)

    # Shown when nothing in the data matches the clicked country
    show_no_studies <- function() {
      showModal(modalDialog(
        title = paste("No studies for", tools::toTitleCase(country_clicked)),
        paste0("There are no studies in this dashboard for ", tools::toTitleCase(country_clicked), "."),
        easyClose = TRUE,
        footer = modalButton("Close")
      ))
    }

    clicked_key <- country_key(country_clicked)
    
    
    country_df <- country_df %>%
      rowwise() %>%
      mutate(Country_click = list(split_country_for_map(Country))) %>%
      ungroup() %>%
      tidyr::unnest_longer(Country_click, values_to = "Country_click") %>%
      mutate(.cn = country_key(Country_click)) %>%
      filter(.cn == clicked_key)
    
    if (nrow(country_df) == 0) {
      show_no_studies()
      return()
    }
    
    display_country <- country_df %>%
      dplyr::pull(Country) %>%
      (\(v) v[!is.na(v) & nzchar(v)])()
    
    display_country <- if (length(display_country) > 0) {
      pretty_country(display_country[1])
    } else {
      pretty_country(country_clicked)
    }
    
    needed_cols <- c(
      "Study.ID", "Author", "Study.Title", "Intervention.Name",
      "Outcome.Domain", "Specific.Outcome", "evidence_rating",
      "publication_year", "Description.of.Intervention",
      paste0("Cleaned.specific.Outcome.", 1:9),
      paste0("Outcome..", 1:9, "..Outcome.effect.category")
    )
    for (nm in needed_cols) {
      if (!nm %in% names(country_df)) country_df[[nm]] <- NA_character_
    }
    
    summary_counts_country <- country_df %>%
      summarise(
        n = n_distinct(Study.ID),
        interventions = {
          nm <- na_if(as.character(Intervention.Name), "")
          idv <- as.character(Intervention.ID)
          n_distinct(idv[!is.na(nm)], nm[!is.na(nm)])
        },
        effective = n_distinct(Study.ID[is_reduced_csv(evidence_rating)]),
        not_effective = n_distinct(Study.ID[is_no_effect_csv(evidence_rating) | is_increased_csv(evidence_rating)]),
        .groups = "drop"
      )
    
    country_card <- function(value, label, accent = sfh_blue) {
      tags$div(
        class = "country-modal-card",
        tags$div(class = "country-modal-accent", style = paste0("background:", accent, ";")),
        tags$div(class = "country-modal-label", label),
        tags$div(class = "country-modal-value", value)
      )
    }
    
    clean_display_text <- function(x) {
      x <- as.character(x)
      x <- enc2utf8(x)
      x <- stringr::str_squish(x)
      x[x == ""] <- NA_character_
      x
    }
    
    collapse_unique_chr <- function(x) {
      x <- clean_display_text(x)
      x <- unique(x[!is.na(x)])
      if (length(x) == 0) return(NA_character_)
      paste(x, collapse = "; ")
    }
    
    max_year_safe <- function(x) {
      x <- suppressWarnings(as.integer(x))
      x <- x[!is.na(x)]
      if (length(x) == 0) return(NA_integer_)
      max(x)
    }
    
    study_display_df <- country_df %>%
      mutate(
        Author = clean_display_text(Author),
        Study.Title = clean_display_text(Study.Title),
        Intervention.Name = clean_display_text(Intervention.Name),
        Outcome.Domain = clean_display_text(Outcome.Domain),
        Specific.Outcome = clean_display_text(Specific.Outcome),
        evidence_rating = clean_display_text(evidence_rating),
        publication_year = suppressWarnings(as.integer(publication_year)),
        `Author (Year)` = dplyr::case_when(
          !is.na(Author) & !is.na(publication_year) & !grepl("\\([0-9]{4}\\)", Author) ~ paste0(Author, " (", publication_year, ")"),
          !is.na(Author) ~ Author,
          !is.na(publication_year) ~ paste0("(", publication_year, ")"),
          TRUE ~ NA_character_
        ),
        `Study title` = Study.Title,
        `Intervention` = Intervention.Name,
        `Outcome domain` = Outcome.Domain,
        `Evidence rating` = evidence_rating,
        `Year` = publication_year
      ) %>%
      distinct()

    full_tbl <- study_display_df %>%
      group_by(`Study title`) %>%
      group_modify(~ tibble(
        `Author (Year)` = collapse_unique_chr(.x$`Author (Year)`),
        `Intervention` = collapse_unique_chr(.x$`Intervention`),
        `Outcome domain` = collapse_unique_chr(.x$`Outcome domain`),
        `Measured CSV outcome(s) and effect` = outcomes_with_effect_html(.x),
        `Evidence rating` = collapse_unique_chr(.x$`Evidence rating`),
        Year = max_year_safe(.x$Year)
      )) %>%
      ungroup() %>%
      arrange(desc(Year), `Study title`)

    html_escape_safe <- function(x) {
      x <- as.character(x)
      not_na <- !is.na(x)
      if (any(not_na)) x[not_na] <- htmltools::htmlEscape(x[not_na])
      x
    }

    full_tbl_display <- full_tbl %>%
      mutate(across(-`Measured CSV outcome(s) and effect`, html_escape_safe))

    full_html <- knitr::kable(
      full_tbl_display,
      format = "html",
      escape = FALSE,
      table.attr = "class='country-modal-table'"
    )
    
    dl_id <- paste0("dl_country_xlsx_", make.names(country_clicked))
    
    showModal(
      modalDialog(
        title = tags$div(
          style = "
    font-family:'GT Pressura Extended', 'General Sans', sans-serif;
    font-size:1.65rem;
    line-height:1.2;
    font-weight:700;
    color:#313131;
  ",
          paste("Country detail \u2014", display_country)
        ),
        size = "l",
        
        fluidRow(
          column(
            8,
            tags$p(
              style = "margin-bottom:12px; font-family:'General Sans', sans-serif; color:#5A5A5A;",
              "The study list for this country is available below."
            )
          ),
          column(12, class = "col-md-4",
                 tags$div(
                   style = "text-align:right; padding-top:2px;",
                   downloadButton(outputId = dl_id, label = "Download country information")
                 )
          )
        ),
        
        fluidRow(
          column(3, country_card(summary_counts_country$n, "Evaluations", sfh_orange)),
          column(3, country_card(summary_counts_country$interventions, "Interventions", sfh_green)),
          column(3, country_card(summary_counts_country$effective, "Effective", sfh_blue)),
          column(3, country_card(summary_counts_country$not_effective, "Not effective", sfh_ocean))
        ),
        
        tags$details(
          style = "margin-top:18px;",
          tags$summary(class = "country-modal-summary", "All studies in this country"),
          tags$div(
            style = "margin-top:10px; overflow-x:auto;",
            HTML(full_html)
          ),
          HTML(outcomes_with_effect_legend_html())
        ),
        
        easyClose = TRUE,
        footer = modalButton("Close")
      )
    )
    
    # Per-country Excel download
    local({
      cty <- country_clicked
      output[[dl_id]] <- downloadHandler(
        filename = function() paste0("CSV_vicperp_", gsub("\\s+", "_", cty), "_summary.xlsx"),
        content = function(file) {
          cty_cn <- country_key(cty)
          dd <- df %>%
            rowwise() %>%
            mutate(Country_click = list(split_country_for_map(Country))) %>%
            ungroup() %>%
            tidyr::unnest_longer(Country_click, values_to = "Country_click") %>%
            mutate(.cn = country_key(Country_click)) %>%
            filter(.cn == cty_cn) %>%
            select(-.cn, -Country_click)   # helper columns, not part of the data
          dd <- dd[, intersect(csv_export_columns, names(dd)), drop = FALSE]
          dd <- clean_export_df(dd)

          write_clean_xlsx(dd, file, "studies")
        }
      )
    })
    
    return()
  }
  
  # ---------------------------------------------------------------------------
  # Filter-chip removal handler (Studies included)
  # ---------------------------------------------------------------------------
  observeEvent(input$remove_filter, {
    val <- input$remove_filter$value
    input_id <- input$remove_filter$input
    
    current <- input[[input_id]]
    
    if (!is.null(current)) {
      new_vals <- setdiff(current, val)
      updateSelectizeInput(
        session,
        input_id,
        selected = new_vals
      )
    }
  })
  
  # ---------------------------------------------------------------------------
  # Dashboard-help and definitions buttons
  # ---------------------------------------------------------------------------
  observeEvent(input$open_dashboard_help,             { showModal(dashboard_help_modal()) })
  observeEvent(input$open_dashboard_help_explorer,    { showModal(dashboard_help_modal()) })
  
  # ---------------------------------------------------------------------------
  # Map click -> country modal
  # ---------------------------------------------------------------------------
  observeEvent(input$world_map_home_shape_click, {
    click <- input$world_map_home_shape_click
    if (is.null(click)) return()
    log_event("map_country_click", click$id)
    map_click_handler(click$id)
  })
  
  # ---------------------------------------------------------------------------
  # Quick views and reset buttons (both tabs)
  # ---------------------------------------------------------------------------
  # Clear every dropdown of a tab, using the filter definitions.
  clear_filters <- function(specs) {
    for (spec in specs) updateSelectizeInput(session, spec$input, selected = character(0))
  }

  # Quick view: pre-fill a dropdown with every choice whose text contains `pattern`.
  quick_pick <- function(input_id, choices, pattern) {
    choices <- unique(choices)
    pick <- choices[str_detect(tolower(trimws(choices)), pattern)]
    updateSelectizeInput(session, input_id, selected = pick)
  }

  level_choices <- unique(unlist(df$Level_List))
  age_choices   <- sort(unique(trimws(unlist(df$Age.Groups_List))))

  # Prevention explorer
  observeEvent(input$int_quick_school,      quick_pick("int_level_filter", level_choices, "school"))
  observeEvent(input$int_quick_community,   quick_pick("int_level_filter", level_choices, "community"))
  observeEvent(input$int_quick_adolescents, quick_pick("int_age_filter", age_choices, "adolesc"))
  observeEvent(input$int_quick_adults,      quick_pick("int_age_filter", age_choices, "adult"))

  # One button per region (created from the data)
  lapply(sort(unique(na.omit(df$World.Region))), function(rg) {
    local({
      region <- rg
      btn_id <- paste0("int_quick_region_", make.names(region))
      observeEvent(input[[btn_id]], {
        updateSelectizeInput(session, "int_region_filter", selected = region)
      }, ignoreInit = TRUE)
    })
  })

  observeEvent(input$int_quick_reset, clear_filters(int_filter_specs))

  # Studies included
  observeEvent(input$quick_school,      quick_pick("context_by_filter", level_choices, "school"))
  observeEvent(input$quick_community,   quick_pick("context_by_filter", level_choices, "community"))
  observeEvent(input$quick_online,      quick_pick("context_by_filter", level_choices, "online"))
  observeEvent(input$quick_adolescents, quick_pick("population_filter", age_choices, "adolesc"))

  # "Reset all filters" (sidebar and home page) clears both tabs
  reset_all_filters <- function() {
    clear_filters(int_filter_specs)
    clear_filters(study_filter_specs)
  }
  observeEvent(input$reset_filters,      reset_all_filters())
  observeEvent(input$reset_filters_home, reset_all_filters())


  # ---------------------------------------------------------------------------
  # Studies included DataTable
  # ---------------------------------------------------------------------------
  
  # Raw column set the table can display (excluding effect / CI columns - filtered below)
  desired_result_cols <- c(
    "Intervention.Name", "Intervention.Length", "Intervention.Length..unit.",
    "Setting", "Description.of.Intervention", "Sample.Size",
    "Recipient.of.Intervention", "Sex", "Age.Range", "Mean.Age",
    "Age..unit.", "Age.Group.Description", "Who.delivered.the.intervention.",
    "Adaptations.for.diverse.groups", "Primary.INSPIRE.Category",
    "Second.INSPIRE.Category", "Study.Design",
    "Methodology", "Primary.or.Secondary.Outcome.", "Outcome.Domain",
    "Specific.Outcome", "Measure", "Reported.by",
    "Summary.of.Key.Findings",
    "Sub.group.analyses.conducted.", "Which.subgroups.",
    "evidence_direction", "evidence_rating", "publication_year",
    "Author", "Study.Title", "Country", "World.Region", "Lastly.updated"
  )
  
  # Display-order priority lists
  priority1 <- c(
    "Action",
    "Study.Title",
    "Intervention.Name",
    "Evaluated.Characteristics",
    "Country",
    "Age.Group.Description",
    "Inspire.Strategies",
    "evidence_rating",
    "Summary.of.Key.Findings"
  )
  priority2 <- c(
    "Study.Design", "Methodology", "Age.Range", "Age.Group.Description",
    "publication_year", "Author"
  )
  priority3 <- c(
    "Sample.Size", "Mean.Age", "Measure",
    "Sub.group.analyses.conducted.", "Which.subgroups."
  )
  
  # Tidy author strings (strip quotes, drop "et al.", trim trailing punctuation)
  clean_author <- function(auth) {
    a <- as.character(auth)
    a[is.na(a)] <- ""
    a <- gsub('\\"', "", a)
    a <- gsub("'", "", a)
    a <- gsub("\\bet al\\.?\\b", "", a, ignore.case = TRUE)
    a <- gsub("[,;:[:space:]]+$", "", a)
    a <- trimws(a)
    a[nzchar(a) == FALSE] <- NA
    a
  }
  
  pretty_name <- function(nm) {
    dplyr::case_when(
      nm == "Study.Title" ~ "Study title",
      nm == "Intervention.Name" ~ "Intervention name",
      nm == "evidence_rating" ~ "Evidence rating",
      nm == "Age.Group.Description" ~ "Age group",
      nm == "publication_year" ~ "Publication year",
      nm == "Study.Design" ~ "Study design",
      nm == "World.Region" ~ "Region",
      nm == "Summary.of.Key.Findings" ~ "Summary of key findings",
      nm == "Inspire.Strategies" ~ "INSPIRE strategies",
      nm == "Primary.INSPIRE.Category" ~ "Primary INSPIRE category",
      nm == "Second.INSPIRE.Category" ~ "Second INSPIRE category",
      TRUE ~ {
        nm2 <- gsub("[._]+", " ", nm)
        nm2 <- trimws(nm2)
        nm2 <- tolower(nm2)
        nm2 <- sub("^([a-z])", "\\U\\1", nm2, perl = TRUE)
        nm2 <- gsub("Inspire", "INSPIRE", nm2)
        nm2
      }
    )
  }
  
  # Compute the displayed column order: priorities first, then the rest
  displayed_cols_reactive <- reactive({
    dcols <- intersect(desired_result_cols, colnames(filtered_data()))
    
    # Exclude effect / CI columns
    exclude_pattern <- "(?i)effect|ci|confidence|odds.ratio|hazard.ratio|beta|se\\b|p.value|pvalue"
    dcols <- dcols[!stringr::str_detect(dcols, exclude_pattern)]
    
    ordered_priorities <- c()
    for (p in priority1) {
      if (p %in% c("Action", "Evaluated.Characteristics", "Inspire.Strategies") || p %in% dcols) {
        ordered_priorities <- c(ordered_priorities, p)
      }
    }
    unique(c(
      ordered_priorities,
      intersect(priority2, dcols),
      intersect(priority3, dcols),
      setdiff(dcols, c(priority1, priority2, priority3))
    ))
  })
  
  # Age-group badge renderer (uses fixed five domains + brand colours)
  make_badges_html <- function(x) {
    if (is.list(x)) x <- unlist(x, use.names = FALSE)
    
    x <- as.character(x)
    x <- x[!is.na(x)]
    x <- trimws(x)
    x <- x[nzchar(x)]
    x <- unique(x)
    
    if (length(x) == 0) return(NA_character_)
    
    # Force consistent display order
    domain_group <- classify_age_group(x)
    ord <- order(match(domain_group, age_group_order), x)
    x <- x[ord]
    domain_group <- domain_group[ord]
    
    # Badge colours (text colour comes from badge_text_colour(); "NA" gets a
    # Sand badge with a Sand-dark outline)
    badge_fill <- c(
      "Early Childhood (0-4 years)"  = sfh_orange,
      "Middle childhood (5-9 years)" = sfh_green,
      "Adolescence (10-18 years)"    = sfh_blue,
      "Adults (18+ years)"           = sfh_purple,
      "NA"                           = sfh_sand
    )

    badges <- vapply(seq_along(x), function(i) {
      label <- x[i]
      grp <- domain_group[i]
      fill <- badge_fill[[grp]]
      border <- if (grp == "NA") sfh_sand_dark else fill
      display_label <- if (grp == "NA") "NA" else label
      
      paste0(
        "<span style='display:inline-block; margin:0; ",
        "padding:3px 8px; border-radius:999px; ",
        "background:", fill, "; border:1px solid ", border, "; ",
        "color:", badge_text_colour(fill), "; font-family:\"General Sans\", sans-serif; ",
        "font-size:11px; line-height:1.25; white-space:nowrap;'>",
        htmltools::htmlEscape(display_label),
        "</span>"
      )
    }, FUN.VALUE = character(1))
    
    paste(badges, collapse = "")
  }
  
  output$results_dt <- renderDT({
    d <- filtered_data()
    d <- d %>% arrange(Intervention.Name, Study.Title)
    d$Intervention.Name[is.na(d$Intervention.Name)] <- "Unknown intervention"
    
    cols_ordered <- displayed_cols_reactive()
    
    # --- Derived display columns ---------------------------------------------
    auth_clean <- clean_author(d$Author)
    has_year_inside <- grepl("\\([0-9]{4}\\)", auth_clean)
    
    eval_char <- ifelse(
      !is.na(auth_clean) & has_year_inside,
      auth_clean,
      ifelse(
        !is.na(auth_clean) & !is.na(d$publication_year),
        paste0(auth_clean, " (", d$publication_year, ")"),
        ifelse(!is.na(auth_clean), auth_clean,
               ifelse(!is.na(d$publication_year),
                      paste0("(", d$publication_year, ")"),
                      NA_character_))
      )
    )
    
    # View button id values
    id_vals <- as.character(d$Study.ID)
    action_col <- sprintf(
      '<button class="btn btn-sm btn-primary view-btn" data-id="%s" title="Open study summary">View</button>',
      htmltools::htmlEscape(id_vals)
    )
    
    # Base table with raw available columns
    raw_cols_available <- intersect(desired_result_cols, colnames(d))
    base_df <- d %>% select(all_of(raw_cols_available))
    
    # Render age-group cells as badges
    base_df$Age.Group.Description <- vapply(
      d$Age.Groups_List,
      make_badges_html,
      FUN.VALUE = character(1)
    )
    
    base_df$Evaluated.Characteristics <- eval_char
    country_display <- as.character(d$Country)
    country_display[country_display %in% c("United States of America", "United States")] <- "USA"
    country_display[country_display %in% c("United Kingdom", "UK", "Great Britain")] <- "UK"
    base_df$Evaluated.In.Country <- country_display

    # INSPIRE strategy badges (Primary + Second INSPIRE category, combined,
    # ordered I-N-S-P-I-R-E), matching the Prevention explorer's treatment.
    base_df$Inspire.Strategies <- vapply(
      d$Inspire_List,
      function(x) make_tag_badges_html(x, order = inspire_category_order),
      FUN.VALUE = character(1)
    )

    # "New" badge based on the Lastly.updated date.
    # A date must be later than the new_badge_cutoff defined above.
    new_flag <- !is.na(d$is_new_update) &
      d$is_new_update

    # Assemble the final display data.frame
    col_map <- list(
      "Evaluated.Characteristics" = "Evaluated.Characteristics",
      "Country" = "Evaluated.In.Country",
      "Inspire.Strategies" = "Inspire.Strategies"
    )
    
    final_col_list <- list()
    for (nm in cols_ordered) {
      if (nm == "Action") {
        final_col_list[["Action"]] <- action_col
      } else if (nm %in% names(col_map)) {
        final_col_list[[nm]] <- base_df[[col_map[[nm]]]]
      } else if (nm %in% colnames(base_df)) {
        final_col_list[[nm]] <- base_df[[nm]]
      } else {
        final_col_list[[nm]] <- rep(NA_character_, nrow(base_df))
      }
    }
    
    new_badges <- ifelse(
      new_flag,
      new_badge_html,
      ""
    )
    
    # Build the data.frame, mapping internal keys to clean output column names
    final_df <- data.frame(New = new_badges, stringsAsFactors = FALSE)
    insert_cols <- list()
    for (nm in names(final_col_list)) {
      out_name <- nm
      if (nm == "Evaluated.Characteristics") {
        out_name <- "Evaluated.Characteristics"
      } else if (nm == "Country") {
        out_name <- "Evaluated.In.Country"
      } else if (nm == "Action") {
        out_name <- "Action"
      }
      insert_cols[[out_name]] <- final_col_list[[nm]]
    }
    for (out_name in names(insert_cols)) {
      final_df[[out_name]] <- insert_cols[[out_name]]
    }
    
    # Truncate very long text fields for display
    final_df <- final_df %>%
      mutate(across(where(is.character),
                    ~ ifelse(nchar(.) > 1000, paste0(substr(., 1, 1000), "\u2026"), .)))
    
    # Pretty header names
    disp_names <- vapply(colnames(final_df), function(x) {
      if (x == "New") return("New")
      if (x == "Evaluated.Characteristics") return("Evaluation author & year")
      if (x == "Evaluated.In.Country") return("Country")
      if (x == "Action") return("View")
      pretty_name(x)
    }, FUN.VALUE = character(1), USE.NAMES = FALSE)
    colnames(final_df) <- disp_names
    
    # Visible-by-default columns and final display order
    desired_order <- c(
      "New",
      "View",
      "Study title",
      "Intervention name",
      "Evidence rating",
      "Country",
      "Evaluation author & year",
      "Age group",
      "INSPIRE strategies"
    )
    other_cols <- setdiff(colnames(final_df), desired_order)
    final_df <- final_df[c(intersect(desired_order, colnames(final_df)), other_cols)]
    
    # Hide everything not in the default-visible list
    hide_pos_1based <- which(!(colnames(final_df) %in% desired_order))
    hide_targets_js <- if (length(hide_pos_1based) > 0) as.integer(hide_pos_1based - 1) else integer(0)
    
    # Make a few long-text columns wider when toggled visible
    wide_cols <- c("Description of intervention", "Summary of key findings", "Study title")
    wide_pos_1based <- which(colnames(final_df) %in% wide_cols)
    wide_targets_js <- if (length(wide_pos_1based) > 0) as.integer(wide_pos_1based - 1) else integer(0)
    
    column_defs <- list(
      list(width = "60px", targets = 0),    # New
      list(width = "80px", targets = 1),    # View
      list(width = "200px", targets = 2),   # Study title
      
      # New and View should not be sortable.
      list(orderable = FALSE, targets = c(0, 1)),
      
      list(className = "dt-head-center", targets = "_all")
    )
    
    if (length(hide_targets_js) > 0) {
      column_defs <- c(column_defs, list(list(visible = FALSE, targets = hide_targets_js)))
    }
    if (length(wide_targets_js) > 0) {
      column_defs <- c(column_defs, list(list(width = "550px", targets = wide_targets_js)))
    }
    
    datatable(
      final_df,
      escape = FALSE,
      rownames = FALSE,
      selection = "none",
      options = list(
        scrollX = TRUE,
        scrollY = "55vh",
        scrollCollapse = TRUE,
        pageLength = 25,
        lengthMenu = list(c(10, 25, 50, 100, -1), c("10", "25", "50", "100", "All")),
        dom = "frtip",
        
        # Initial table order:
        # Intervention name A-Z, then Study title A-Z.
        # Unlike orderFixed, users can replace this by clicking another heading.
        order = list(
          list(3, "asc"),
          list(2, "asc")
        ),
        
        ordering = TRUE,
        columnDefs = column_defs
      )
    )
  }, server = TRUE)
  
  
  # ---------------------------------------------------------------------------
  # Download full CSV (Home)
  # ---------------------------------------------------------------------------
  observeEvent(input$download_summary_all, {
    showModal(
      modalDialog(
        title = "Download",
        "Click the button to download a CSV file of the dataset with any applied filters. If you would like to download information about all studies, please clear filters before downloading.",
        footer = tagList(
          downloadButton("download_summary_handler", "Download full information"),
          modalButton("Close")
        )
      )
    )
  })
  
  output$download_summary_handler <- downloadHandler(
    filename = function() paste0("CSV_vicperp_evidence_summary_", Sys.Date(), ".xlsx"),
    content = function(file) {
      dd <- normalise_df(filtered_data())
      dd <- dd[, intersect(csv_export_columns, names(dd)), drop = FALSE]
      dd <- clean_export_df(dd)
      write_clean_xlsx(dd, file, "evidence_summary")
      log_event("download", "download_summary_filtered_xlsx")
    }
  )
  
  # ---------------------------------------------------------------------------
  # Cite-dashboard modal + copy-to-clipboard
  # ---------------------------------------------------------------------------
  observeEvent(input$cite_dashboard, {
    showModal(
      modalDialog(
        title = "Cite this dashboard",
        size = "m",
        tags$p("Copy the citation below for use in reports or publications:"),
        tags$div(
          style = "margin-top:8px;",
          tags$textarea(
            id = "citation_box",
            rows = 4,
            style = "width:100%; font-family:'General Sans', sans-serif;",
            citation_text_dashboard
          )
        ),
        footer = tagList(
          actionButton("copy_citation_btn", "Copy to clipboard"),
          modalButton("Close")
        ),
        easyClose = TRUE
      )
    )
  })
  
  observeEvent(input$copy_citation_btn, {
    session$sendCustomMessage(
      "copy-citation-to-clipboard",
      list(text = citation_text_dashboard)
    )
  })
  
  observeEvent(input$citation_copy_status, {
    if (input$citation_copy_status %in% c("copied", "copied_fallback")) {
      showNotification("Citation copied to clipboard.", type = "message", duration = 3)
    } else {
      showNotification(
        "Could not copy automatically. Please select the citation text and copy it manually.",
        type = "warning",
        duration = 5
      )
    }
  })
  
  # ---------------------------------------------------------------------------
  # Feedback + Share-evidence modals (embedded Microsoft Forms)
  # ---------------------------------------------------------------------------
  observeEvent(input$open_feedback_form, {
    showModal(
      modalDialog(
        title = "Feedback",
        size = "l",
        tags$iframe(
          src                  = "https://forms.cloud.microsoft/e/2bVT1QywKg?embed=true",
          width                = "100%",
          height               = "480px",
          frameborder          = "0",
          marginwidth          = "0",
          marginheight         = "0",
          style                = "border:none; max-width:100%; max-height:100vh;",
          allowfullscreen      = NA,
          webkitallowfullscreen = NA,
          mozallowfullscreen   = NA,
          msallowfullscreen    = NA
        ),
        easyClose = TRUE,
        footer = modalButton("Close")
      )
    )
  })
  
  observeEvent(input$open_share_evidence_form, {
    showModal(
      modalDialog(
        title = "Share evidence or interventions",
        size = "l",
        tags$iframe(
          src                  = "https://forms.cloud.microsoft/e/3iKNjknJcZ?embed=true",
          width                = "100%",
          height               = "480px",
          frameborder          = "0",
          marginwidth          = "0",
          marginheight         = "0",
          style                = "border:none; max-width:100%; max-height:100vh;",
          allowfullscreen      = NA,
          webkitallowfullscreen = NA,
          mozallowfullscreen   = NA,
          msallowfullscreen    = NA
        ),
        easyClose = TRUE,
        footer = modalButton("Close")
      )
    )
  })
  
  
  # ---------------------------------------------------------------------------
  # Helpers for the study / intervention pop-up
  # ---------------------------------------------------------------------------
  first_nonmissing <- function(x) {
    x <- as.character(x)
    x <- x[!is.na(x)]
    x <- trimws(x)
    x <- x[nzchar(x)]
    if (length(x) == 0) return(NA_character_)
    x[1]
  }
  
  # Shared "outcome name (domain) + effect pill" rendering (down arrow =
  # Reduced, straight line = No change, up arrow = Increased), matching the
  # Prevention explorer's "Underlying evaluations since 2018" table/legend.
  # There is no "Mixed" category: each outcome (domain) shows a single pill,
  # resolved by resolve_outcome_effect() (any increase = I, otherwise any
  # reduction = R, otherwise NC).
  outcome_effect_pill_colours <- c(R = sfh_green, NC = sfh_charcoal, I = sfh_pink)
  outcome_effect_pill_symbols <- c(R = "↓", NC = "—", I = "↑")
  outcome_effect_pill_titles  <- c(R = "Reduced CSV", NC = "No change in CSV", I = "Increased CSV")

  outcome_effect_pill <- function(codes) {
    cd <- resolve_outcome_effect(codes)
    if (is.na(cd)) return("")

    paste0(
      "<span style='display:inline-block; margin-left:5px; padding:1px 7px; ",
      "border-radius:999px; background:", outcome_effect_pill_colours[[cd]], "; color:", badge_text_colour(outcome_effect_pill_colours[[cd]]), "; ",
      "font-family:\"General Sans\"; font-size:10px; font-weight:700; ",
      "white-space:nowrap; vertical-align:middle;' title='",
      outcome_effect_pill_titles[[cd]], "'>", outcome_effect_pill_symbols[[cd]], "</span>"
    )
  }

  # Build "outcome (domain) <pill>" lines (one per line) across every row of d
  # (a study's, or a country's-worth of, evaluation rows); see
  # outcome_effect_lines() for how repeat measurements are collapsed.
  outcomes_with_effect_html <- function(d) {
    if (is.null(d) || nrow(d) == 0) return("Not reported")

    lines <- outcome_effect_lines(d)
    if (nrow(lines) == 0) return("Not reported")

    rows <- vapply(seq_len(nrow(lines)), function(i) {
      paste0(
        "<div style='margin-bottom:3px;'>",
        htmltools::htmlEscape(lines$label[i]),
        outcome_effect_pill(lines$code[i]),
        "</div>"
      )
    }, FUN.VALUE = character(1))

    paste(rows, collapse = "")
  }

  # Legend for the outcome-effect pills, shown once at the bottom of any
  # table/box that uses outcomes_with_effect_html()/outcome_effect_pill().
  #
  # Styled identically to the Prevention explorer's legend (solid pills, same
  # as the real ratings). In the study pop-out it is placed OUTSIDE (directly
  # under) the green outcomes box.
  outcomes_with_effect_legend_html <- function() {
    paste0(
      "<div style='margin-top:8px; font-size:11px; color:", sfh_charcoal_mid, "; line-height:1.6;'>",
      "<span style='display:inline-block; padding:1px 7px; border-radius:999px; background:", sfh_green, "; color:", badge_text_colour(sfh_green), "; font-weight:700; margin-right:4px;'>↓</span>",
      "Reduced CSV",
      "<span style='display:inline-block; padding:1px 7px; border-radius:999px; background:", sfh_charcoal, "; color:", badge_text_colour(sfh_charcoal), "; font-weight:700; margin:0 4px 0 8px;'>—</span>",
      "No change in CSV",
      "<span style='display:inline-block; padding:1px 7px; border-radius:999px; background:", sfh_pink, "; color:", badge_text_colour(sfh_pink), "; font-weight:700; margin:0 4px 0 8px;'>↑</span>",
      "Increased CSV",
      "</div>"
    )
  }

  overall_csv_effect_sentence <- function(effect_value) {
    effect_value <- first_nonmissing(effect_value)
    
    if (is.na(effect_value)) {
      return("This study was not categorised for overall CSV effectiveness.")
    }
    
    label <- get_evidence_rating_v3(effect_value)
    
    dplyr::case_when(
      label == "Reduced CSV" ~
        "Across the CSV outcomes evaluated, this study found that the intervention reduced CSV.",
      label == "No change in CSV" ~
        "Across the CSV outcomes evaluated, this study found no effect on CSV.",
      label == "Increased CSV" ~
        "Across the CSV outcomes evaluated, this study found that the intervention was associated with increased CSV.",
      TRUE ~
        "This study was not categorised for overall CSV effectiveness."
    )
  }
  
  make_study_summary_box <- function(title, body, accent = sfh_green) {
    tags$div(
      style = paste0(
        "margin-top:15px; padding:15px; background-color:", sfh_sand, "; ",
        "border-left:4px solid ", accent, "; border-radius:var(--sfh-panel-radius); ",
        "border-top:1px solid rgba(49,49,49,0.08); ",
        "border-right:1px solid rgba(49,49,49,0.08); ",
        "border-bottom:1px solid rgba(49,49,49,0.08);"
      ),
      tags$b(
        style = paste0("color:", accent, "; font-family:'GT Pressura Extended', sans-serif;"),
        title
      ),
      tags$div(
        style = paste0("margin-top:8px; color:", sfh_charcoal, "; font-family:'General Sans', sans-serif; line-height:1.5;"),
        body
      )
    )
  }
  
  
  # ---------------------------------------------------------------------------
  # Study / Intervention pop-up (triggered by the table "View" buttons)
  # ---------------------------------------------------------------------------
  observeEvent(input$show_study_id, {
    req(input$show_study_id)
    sid <- as.character(input$show_study_id)
    log_event("view_open", sid)
    
    # Local convenience helpers (used in both branches)

    # Cleans a raw DOI field (which may already be a full URL, a "doi:"
    # prefix, or a bare DOI) into a canonical https://doi.org/... link.
    # Returns NA_character_ when there is no usable DOI.
    clean_doi_url_local <- function(doi) {
      doi <- as.character(doi)
      doi <- trimws(doi)
      if (is.na(doi) || !nzchar(doi)) return(NA_character_)

      doi_clean <- doi
      doi_clean <- sub("^https?://(www\\.)?doi\\.org/", "", doi_clean, ignore.case = TRUE)
      doi_clean <- sub("^(www\\.)?doi\\.org/", "", doi_clean, ignore.case = TRUE)
      doi_clean <- sub("^doi:\\s*", "", doi_clean, ignore.case = TRUE)
      doi_clean <- trimws(doi_clean)
      doi_clean <- sub("^(www\\.)?doi\\.org/", "", doi_clean, ignore.case = TRUE)
      doi_clean <- trimws(doi_clean)

      if (!nzchar(doi_clean) || tolower(doi_clean) %in% c("na", "n/a", "missing")) {
        return(NA_character_)
      }

      paste0("https://doi.org/", doi_clean)
    }

    # htmltools::htmlEscape() errors on vectors that mix NA with
    # otherwise-unescaped text (any(grepl(...)) can evaluate to NA).
    # This wrapper escapes only the non-NA elements and leaves NA as NA.
    html_escape_safe <- function(x) {
      x <- as.character(x)
      not_na <- !is.na(x)
      if (any(not_na)) x[not_na] <- htmltools::htmlEscape(x[not_na])
      x
    }

    collapse_unique_text <- function(x) {
      x <- as.character(x)
      x <- trimws(x)
      x <- x[!is.na(x) & nzchar(x)]
      x <- unique(x)
      if (length(x) == 0) return(NA_character_)
      paste(x, collapse = ", ")
    }
    
    normalise_subgroup_conducted <- function(x) {
      x <- as.character(x)
      x <- x[!is.na(x)]
      x <- stringr::str_squish(x)
      x <- x[nzchar(x)]
      if (length(x) == 0) return("Not reported")
      txt <- tolower(paste(x, collapse = "; "))
      if (grepl("^yes\\b", txt)) {
        "Yes"
      } else if (grepl("^no\\b", txt)) {
        "No"
      } else {
        "Not reported"
      }
    }
    
    clean_subgroup_label <- function(x) {
      x <- as.character(x)
      x <- x[!is.na(x)]
      x <- stringr::str_squish(x)
      x <- x[nzchar(x)]
      if (length(x) == 0) return(character(0))
      
      parts <- unlist(strsplit(paste(x, collapse = "; "), ";", fixed = TRUE))
      parts <- stringr::str_squish(parts)
      parts <- parts[nzchar(parts)]
      
      # Strip explanatory text
      parts <- sub("\\s*\\[.*$", "", parts)
      parts <- sub("\\s*\\(.*$", "", parts)
      parts <- sub("\\s+This intervention.*$", "", parts, ignore.case = TRUE)
      parts <- sub("\\s+The intervention.*$", "", parts, ignore.case = TRUE)
      parts <- sub("\\s+No significant.*$", "", parts, ignore.case = TRUE)
      parts <- sub("\\s+Significant effects.*$", "", parts, ignore.case = TRUE)
      
      parts <- stringr::str_squish(parts)
      parts <- parts[nzchar(parts)]
      
      # Tokenise then extract known subgroup categories cleanly
      tokens <- unlist(strsplit(parts, "\\s+"))
      tokens <- stringr::str_squish(tokens)
      tokens <- tokens[nzchar(tokens)]
      tokens <- paste(tokens, collapse = " ")
      
      groups <- c()
      if (grepl("\\bage\\b", tokens, ignore.case = TRUE)) groups <- c(groups, "Age")
      if (grepl("\\bsex\\b", tokens, ignore.case = TRUE)) groups <- c(groups, "Sex")
      if (grepl("\\brace\\b", tokens, ignore.case = TRUE)) groups <- c(groups, "Race/Ethnicity")
      if (grepl("\\bethnicity\\b", tokens, ignore.case = TRUE)) groups <- c(groups, "Race/Ethnicity")
      if (grepl("sexual orientation", tokens, ignore.case = TRUE)) groups <- c(groups, "Sexual orientation")
      
      if (length(groups) == 0) groups <- parts
      unique(groups)
    }
    
    subgroup_display_text <- function(conducted, which_subgroups) {
      conducted_clean <- normalise_subgroup_conducted(conducted)
      which_txt <- clean_subgroup_label(which_subgroups)
      which_txt <- which_txt[!tolower(which_txt) %in% c("n/a", "na", "not reported", "not available", "missing")]
      if (conducted_clean == "Yes" && length(which_txt) > 0) {
        paste0("Yes \u2014 ", paste(which_txt, collapse = "; "))
      } else {
        conducted_clean
      }
    }
    
    # ------------------------------------------------------------------------
    # Branch 1: Prevention explorer pop-up
    # ------------------------------------------------------------------------
    if (grepl("^intid::|^intname::", sid)) {
      
      is_int_id <- grepl("^intid::", sid)
      int_id   <- if (is_int_id) sub("^intid::", "", sid) else NA_character_
      int_name <- if (!is_int_id) URLdecode(sub("^intname::", "", sid)) else NA_character_
      
      int_id   <- trimws(as.character(int_id))
      int_name <- trimws(as.character(int_name))
      
      # The evaluation rows (effective ones only) that belong to this intervention
      rows_for_intervention <- function(src) {
        src %>%
          mutate(
            Intervention.ID = as.character(Intervention.ID),
            Intervention.Name = as.character(Intervention.Name),
            intervention_effectiveness_code = toupper(trimws(as.character(intervention_effectiveness_code)))
          ) %>%
          filter(
            intervention_effectiveness_code == "R",
            if (is_int_id) {
              !is.na(Intervention.ID) & trimws(Intervention.ID) == int_id
            } else {
              !is.na(Intervention.Name) & trimws(Intervention.Name) == int_name
            }
          )
      }

      d <- rows_for_intervention(filtered_data())

      # If filters hid the intervention, fall back to the full df
      if (nrow(d) == 0) d <- rows_for_intervention(df)
      
      if (nrow(d) == 0) {
        showModal(
          modalDialog(
            title = "Intervention not found",
            "Could not find the selected intervention in the current filtered results.",
            easyClose = TRUE,
            footer = modalButton("Close")
          )
        )
        return()
      }
      
      n_evals <- d %>% filter(!is.na(Study.ID)) %>% distinct(Study.ID) %>% nrow()
      
      countries_vec <- if ("Country" %in% names(d)) {
        d %>%
          pull(Country) %>%
          as.character() %>%
          trimws() %>%
          (\(x) x[!is.na(x) & nzchar(x)])() %>%
          (\(x) unique(unlist(lapply(x, split_country_tokens))))()
      } else {
        character(0)
      }
      
      n_countries <- length(countries_vec)
      countries_text <- if (length(countries_vec) > 0) {
        paste(countries_vec, collapse = ", ")
      } else {
        "Not reported"
      }
      
      # CSV outcome domains
      outcomes <- unique(trimws(unlist(d$Outcome_Domain_List)))
      outcomes <- outcomes[!is.na(outcomes) & nzchar(outcomes)]
      
      outcomes_html <- if (length(outcomes) > 0) {
        make_tag_badges_html(
          outcomes,
          order = c("Victimisation", "Perpetration",
                    "Both (combined vic/perp outcome)"),
          special_fill = c("Both (combined vic/perp outcome)" = sfh_pink)
        )
      } else {
        "Not reported"
      }
      
      # Implementation contexts
      locations_vec <- unique(trimws(unlist(d$Level_List)))
      locations_vec <- locations_vec[!is.na(locations_vec) & nzchar(locations_vec)]
      
      locations_html <- if (length(locations_vec) > 0) {
        make_tag_badges_html(
          locations_vec,
          order = context_order
        )
      } else {
        "Not reported"
      }
      locations_text <- if (length(locations_vec) > 0) {
        paste(locations_vec, collapse = ", ")
      } else {
        "Not reported"
      }
      
      # Intervention description details - Target / Design composite shown
      # under "Intervention description" in this pop-up's details section.
      population_target <- if ("Intervention.Target" %in% names(d)) {
        pick_richest_value(d$Intervention.Target)
      } else {
        NA_character_
      }
      
      population_design <- if ("Study.Design" %in% names(d)) {
        first_nonmissing(d$Study.Design)
      } else {
        NA_character_
      }
      population_target   <- if (is.na(population_target) || !nzchar(population_target)) "Not reported" else population_target
      population_design   <- if (is.na(population_design) || !nzchar(population_design)) "Not reported" else population_design
      
      # No "Context" row here - it's the same Level.location.of.intervention
      # field as "Contexts where the intervention was implemented" directly
      # below, so showing it twice back-to-back would be redundant.
      population_html <- paste0(
        "<div style='display:flex; flex-direction:column; gap:4px;'>",
        "<div><strong>Type of CSV victimisation or perpetration outcome(s):</strong> ", make_uniform_badge(population_target, sfh_green), "</div>",
        "<div><strong>Design:</strong> ", make_uniform_badge(population_design, sfh_blue), "</div>",
        "</div>"
      )
      
      int_desc <- if ("Description.of.Intervention" %in% names(d)) {
        first_nonmissing(d$Description.of.Intervention)
      } else {
        NA_character_
      }
      if (is.na(int_desc)) int_desc <- "No description available."
      
      cleaned_outcome_by_row <- vapply(
        seq_len(nrow(d)),
        function(i) outcomes_with_effect_html(d[i, , drop = FALSE]),
        FUN.VALUE = character(1)
      )
      
      # Link each row's Author (Year) text to its DOI, when available,
      # matching the Studies included tab's DOI-linking behaviour.
      doi_col_int <- grep("^DOI$|^Doi$|doi", names(d), value = TRUE, ignore.case = TRUE)[1]
      doi_url_by_row <- if (!is.na(doi_col_int) && doi_col_int %in% names(d)) {
        vapply(d[[doi_col_int]], clean_doi_url_local, FUN.VALUE = character(1))
      } else {
        rep(NA_character_, nrow(d))
      }
      author_text_esc <- html_escape_safe(d$Author)
      author_year_html_by_row <- vapply(seq_len(nrow(d)), function(i) {
        au <- author_text_esc[i]
        if (is.na(au)) au <- ""
        du <- doi_url_by_row[i]
        if (!is.na(du)) {
          paste0(
            "<a href='", htmltools::htmlEscape(du, attribute = TRUE),
            "' target='_blank' rel='noopener noreferrer'>", au, "</a>"
          )
        } else {
          au
        }
      }, FUN.VALUE = character(1))

      studies_tbl <- d %>%
        mutate(
          `Author (Year)` = author_year_html_by_row,
          Sex = html_escape_safe(Sex),
          `Age Group` = html_escape_safe(if ("Age.Group.Description" %in% names(.)) {
            sapply(Age.Group.Description, clean_age_groups)
          } else {
            NA_character_
          }),
          `Measured CSV outcome(s) and effect` = cleaned_outcome_by_row,
          `INSPIRE strategy` = html_escape_safe({
            prim <- clean_inspire_category(Primary.INSPIRE.Category)
            sec  <- clean_inspire_category(Second.INSPIRE.Category)
            mapply(function(a, b) make_tag_plain(c(a, b), order = inspire_category_order), prim, sec, USE.NAMES = FALSE)
          })
        ) %>%
        select(
          `Author (Year)`,
          Sex,
          `Age Group`,
          `Measured CSV outcome(s) and effect`,
          `INSPIRE strategy`
        ) %>%
        distinct()

      studies_html <- knitr::kable(
        studies_tbl,
        format = "html",
        escape = FALSE,
        table.attr = "class='table table-condensed' style='font-size:13px;'"
      )
      
      overview_card <- function(title, value, subtitle = NULL,
                                accent = sfh_blue, body_html = NULL) {
        tags$div(
          style = paste0(
            "flex:1 1 190px; min-width:190px; padding:12px 14px; ",
            "background:#FFFFFF; border:1px solid rgba(49,49,49,0.10); ",
            "border-top:4px solid ", accent, "; border-radius:var(--sfh-panel-radius);"
          ),
          tags$div(
            style = paste0("font-family:'General Sans', sans-serif; font-size:12px; color:", sfh_charcoal_mid, "; margin-bottom:5px;"),
            title
          ),
          tags$div(
            style = paste0("font-family:'GT Pressura Extended', sans-serif; font-size:28px; font-weight:700; color:", sfh_charcoal, "; line-height:1;"),
            value
          ),
          if (!is.null(subtitle)) {
            tags$div(
              style = paste0("font-family:'General Sans', sans-serif; font-size:12px; color:", sfh_charcoal_mid, "; margin-top:6px; line-height:1.35;"),
              subtitle
            )
          },
          if (!is.null(body_html)) {
            tags$div(
              style = "margin-top:7px; font-family:'General Sans', sans-serif; font-size:12px; line-height:1.35;",
              HTML(body_html)
            )
          }
        )
      }
      
      showModal(
        modalDialog(
          title = modal_title(paste("Intervention summary -", first_nonmissing(d$Intervention.Name))),
          size = "l",
          fluidRow(
            column(
              12,
              
              tags$div(
                style = paste0("margin-top:8px; margin-bottom:14px; padding:14px 16px; background:", sfh_sand, "; border:1px solid rgba(49,49,49,0.10); border-radius:var(--sfh-card-radius);"),
                tags$div(
                  style = paste0("font-family:'GT Pressura Extended', sans-serif; color:", sfh_ocean, "; font-size:14px; font-weight:700; margin-bottom:10px;"),
                  "Overview"
                ),
                tags$div(
                  style = "display:flex; gap:10px; flex-wrap:wrap;",
                  overview_card(
                    title = "Evaluations",
                    value = n_evals,
                    accent = sfh_orange
                  ),
                  overview_card(
                    title = "Countries",
                    value = n_countries,
                    subtitle = countries_text,
                    accent = sfh_green
                  ),
                  overview_card(
                    title = "CSV victimisation and perpetration outcome(s)",
                    value = length(outcomes),
                    body_html = outcomes_html,
                    accent = sfh_blue
                  ),
                  overview_card(
                    title = "Location(s) of implementation",
                    value = length(locations_vec),
                    body_html = locations_html,
                    accent = sfh_purple
                  )
                )
              ),
              
              tags$div(
                style = paste0(
                  "margin-top:20px; padding:15px; background-color:", sfh_sand, "; ",
                  "border-left:4px solid ", sfh_purple, "; border-radius:12px;"
                ),
                tags$b(style = paste0("color:", sfh_purple, ";"), "Description"),
                tags$p(style = paste0("margin-top:5px; color:", sfh_charcoal, ";"), int_desc)
              ),
              
              tags$details(
                style = "margin-top:12px;",
                tags$summary(
                  style = paste0("cursor:pointer; font-family:'GT Pressura Extended', sans-serif; color:", sfh_ocean, "; font-size:14px;"),
                  "Intervention details"
                ),
                tags$div(
                  style = paste0("margin-top:10px; padding:12px; background:", sfh_sand, "; border:1px solid rgba(49,49,49,0.10); border-radius:var(--sfh-panel-radius);"),
                  tags$div(
                    style = "margin-bottom:12px;",
                    tags$strong("Intervention description"),
                    tags$div(style = "margin-top:4px;", HTML(population_html))
                  ),
                  tags$p(
                    tags$strong("Contexts where the intervention was implemented"),
                    tags$br(),
                    locations_text
                  ),
                  tags$p(
                    tags$strong("Countries where the intervention was implemented and evaluated"),
                    tags$br(),
                    countries_text
                  )
                )
              ),
              
              hr(),
              
              tags$b("Underlying evaluations since 2018"),
              tags$div(
                style = paste0(
                  "max-height:400px; max-width:100%; width:100%; ",
                  "overflow-x:auto; overflow-y:auto; margin-top:10px; ",
                  "box-sizing:border-box;"
                ),
                HTML(studies_html)
              ),
              HTML(outcomes_with_effect_legend_html())
            )
          ),
          easyClose = TRUE,
          footer = modalButton("Close")
        )
      )
      return()
    }
    
    
    # ------------------------------------------------------------------------
    # Branch 2: Studies included study pop-up
    # ------------------------------------------------------------------------
    rows_for_study <- function(src) {
      src %>%
        mutate(Study.ID = as.character(Study.ID)) %>%
        filter(Study.ID == sid)
    }

    study_row <- rows_for_study(filtered_data())

    # If filters hid the study, fall back to the full df
    if (nrow(study_row) == 0) study_row <- rows_for_study(df)
    
    if (nrow(study_row) == 0) {
      showModal(
        modalDialog(
          title = "Study not found",
          "Could not find the selected study.",
          easyClose = TRUE,
          footer = modalButton("Close")
        )
      )
      return()
    }
    
    # All rows belonging to this study (to capture all outcomes)
    study_all_rows <- rows_for_study(df)
    
    study_row <- study_row[1, , drop = FALSE]
    
    study_title <- first_nonmissing(study_row$Study.Title)
    if (is.na(study_title)) study_title <- first_nonmissing(study_row$Title)
    if (is.na(study_title)) study_title <- "Study summary"

    # Filesystem-safe version of the study's short title, for the download
    # filename - falls back to the full title if there is no short title.
    # When the short title is shared by another study, a word-bounded slice
    # of the full title (<=40 chars) is appended so the two downloads never
    # suggest the same file name, without making it unreadably long.
    download_short_title <- first_nonmissing(study_row$Short.Title)
    if (is.na(download_short_title) || !nzchar(trimws(download_short_title))) {
      download_short_title <- clean_filename_part(study_title)
    } else if (trimws(download_short_title) %in% duplicated_short_titles) {
      title_slice <- substr(study_title, 1, 40)
      title_slice <- sub("\\s+\\S*$", "", title_slice)
      if (!nzchar(title_slice)) title_slice <- substr(study_title, 1, 40)
      download_short_title <- paste0(clean_filename_part(download_short_title), "_", clean_filename_part(title_slice))
    } else {
      download_short_title <- clean_filename_part(download_short_title)
    }
    if (!nzchar(download_short_title)) download_short_title <- sid
    
    author_text       <- first_nonmissing(study_row$Author)
    # Drop a year already embedded in the raw Author text (e.g. "Waterman
    # (2022)") since it's shown again right below as its own "Publication
    # year" line - showing it twice reads as repetitive.
    if (!is.na(author_text)) author_text <- trimws(gsub("\\s*\\([0-9]{4}\\)\\s*$", "", author_text))
    pub_year_text     <- first_nonmissing(study_row$publication_year)
    if (is.na(pub_year_text)) pub_year_text <- first_nonmissing(study_row$Year)
    country_text      <- first_nonmissing(study_row$Country)
    intervention_text <- first_nonmissing(study_row$Intervention.Name)
    
    doi_col <- grep("^DOI$|^Doi$|doi", names(study_row), value = TRUE, ignore.case = TRUE)[1]
    doi_text <- if (!is.na(doi_col)) {
      first_nonmissing(study_row[[doi_col]])
    } else {
      NA_character_
    }
    
    make_doi_link <- function(doi) {
      doi_url <- clean_doi_url_local(doi)
      if (is.na(doi_url)) return("")

      paste0(
        "<li><strong>DOI:</strong> ",
        "<a href='", htmltools::htmlEscape(doi_url, attribute = TRUE),
        "' target='_blank' rel='noopener noreferrer'>",
        htmltools::htmlEscape(sub("^https://doi\\.org/", "", doi_url)),
        "</a></li>"
      )
    }
    
    # One outcome per (cleaned specific outcome, domain) line listed below.
    n_outcomes <- nrow(outcome_effect_lines(study_all_rows))
    outcomes_effect_html <- outcomes_with_effect_html(study_all_rows)

    effect_value <- if ("Effectiveness.v3...Effective." %in% names(study_row)) {
      study_row$Effectiveness.v3...Effective.
    } else {
      NA_character_
    }
    effect_sentence <- overall_csv_effect_sentence(effect_value)
    
    description_text <- first_nonmissing(study_row$Description.of.Intervention)
    if (is.na(description_text)) description_text <- "No description available."
    
    intervention_statement <- if (!is.na(intervention_text) && nzchar(intervention_text)) {
      paste0("This study evaluated ", intervention_text, ".")
    } else {
      "This study evaluated an intervention that was not named in the extracted data."
    }
    
    # Coloured summary boxes (intro / description / outcomes / effect)
    study_intro_box <- make_study_summary_box(
      title = "Evaluated intervention",
      body = tags$p(style = "margin:0;", intervention_statement),
      accent = sfh_orange
    )
    study_description_box <- make_study_summary_box(
      title = "Intervention description",
      body = tags$p(style = "margin:0;", description_text),
      accent = sfh_purple
    )
    study_outcomes_box <- make_study_summary_box(
      title = "CSV victimisation or perpetration outcomes evaluated and effect",
      body = tagList(
        tags$p(
          style = "margin:0 0 6px 0;",
          paste0(
            "This study evaluated the effect on ",
            n_outcomes,
            " CSV outcome",
            ifelse(n_outcomes == 1, "", "s"),
            ":"
          )
        ),
        HTML(outcomes_effect_html)
      ),
      accent = sfh_green
    )
    study_effect_box <- make_study_summary_box(
      title = "Overall effect rating",
      body = tags$p(style = "margin:0;", effect_sentence),
      accent = sfh_blue
    )
    
    dl_study_id <- paste0("download_study_", gsub("[^[:alnum:]]", "_", sid))
    
    showModal(
      modalDialog(
        title = modal_title(paste("Study summary \u2014", study_title)),
        size = "l",
        fluidRow(
          column(
            8,
            study_intro_box,
            study_description_box,
            study_outcomes_box,
            HTML(outcomes_with_effect_legend_html()),
            study_effect_box,
            
            tags$p(
              style = paste0("margin-top:15px; color:", sfh_charcoal_mid, ";"),
              "Explore more information about the evaluation and the intervention below by clicking on the dropdowns or by downloading the study as an Excel file."
            ),
            
            tags$details(
              style = "margin-top:12px;",
              tags$summary(
                style = paste0("cursor:pointer; font-family:'GT Pressura Extended', sans-serif; color:", sfh_ocean, "; font-size:14px;"),
                "More study information"
              ),
              tags$div(
                style = paste0("margin-top:10px; padding:12px; background:", sfh_sand, "; border:1px solid rgba(49,49,49,0.10); border-radius:var(--sfh-panel-radius);"),
                tags$p(tags$strong("Study title"), tags$br(), first_nonmissing(study_row$Study.Title)),
                tags$p(tags$strong("Study design"), tags$br(), study_design_tag(first_nonmissing(study_row$Study.Design))),
                tags$p(tags$strong("Sample size"), tags$br(), first_nonmissing(study_row$Sample.Size)),
                tags$p(
                  tags$strong("Subgroup analyses conducted"),
                  tags$br(),
                  subgroup_display_text(
                    study_row$Sub.group.analyses.conducted.,
                    study_row$Which.subgroups.
                  )
                )
              )
            ),
            
            tags$details(
              style = "margin-top:12px;",
              tags$summary(
                style = paste0("cursor:pointer; font-family:'GT Pressura Extended', sans-serif; color:", sfh_ocean, "; font-size:14px;"),
                "More intervention information"
              ),
              tags$div(
                style = paste0("margin-top:10px; padding:12px; background:", sfh_sand, "; border:1px solid rgba(49,49,49,0.10); border-radius:var(--sfh-panel-radius);"),
                tags$p(tags$strong("Sex of study participants"), tags$br(), first_nonmissing(study_row$Sex)),
                tags$p(tags$strong("Age group of study participants"), tags$br(), first_nonmissing(study_row$Age.Group.Description)),
                tags$p(
                  tags$strong("Context where the intervention was implemented"),
                  tags$br(),
                  collapse_unique_text(unlist(study_row$Level_List))
                )
              )
            )
          ),
          
          column(12, class = "col-md-4",
                 tags$b("Key details"),
                 HTML(
                   paste0(
                     "<ul>",
                     if (!is.na(author_text)) paste0("<li><strong>Author:</strong> ", htmltools::htmlEscape(author_text), "</li>") else "",
                     if (!is.na(pub_year_text))    paste0("<li><strong>Publication year:</strong> ", htmltools::htmlEscape(pub_year_text), "</li>") else "",
                     if (!is.na(country_text))     paste0("<li><strong>Country study was implemented:</strong> ", htmltools::htmlEscape(country_text), "</li>") else "",
                     make_doi_link(doi_text),
                     "</ul>"
                   )
                 ),
                 br(),
                 downloadButton(outputId = dl_study_id, label = "Download study information")
          )
        ),
        easyClose = TRUE,
        footer = modalButton("Close")
      )
    )
    
    # Per-study Excel download
    local({
      this_dl_study_id <- dl_study_id
      this_sid <- sid
      this_short_title <- download_short_title

      output[[this_dl_study_id]] <- downloadHandler(
        filename = function() {
          paste0("CSV_vicperp_study_", this_short_title, ".xlsx")
        },
        content = function(file) {
          dd <- df
          
          dd <- dd %>%
            mutate(Study.ID = as.character(Study.ID)) %>%
            filter(Study.ID == this_sid)
          dd <- dd[, intersect(csv_export_columns, names(dd)), drop = FALSE]
          dd <- clean_export_df(dd)

          write_clean_xlsx(dd, file, "study")
        }
      )
    })
  })
  
  
  # ---------------------------------------------------------------------------
  # Welcome / Help / Definitions modal builders
  # ---------------------------------------------------------------------------
  
  dashboard_help_modal <- function() {
    modalDialog(
      title = modal_title("How to use this dashboard"),
      easyClose = TRUE,
      size = "l",
      footer = modalButton("Close"),
      
      tags$div(
        style = "line-height:1.6; font-family:'General Sans', sans-serif;",
        
        modal_heading("The dashboard contains two tabs: Prevention explorer and Studies included."),
        
        modal_heading("What you can do in each tab", margin_top = "22px", margin_bottom = "12px"),
        
        tags$p(
          tags$strong("Prevention explorer: "),
          "Use this tab to explore effective interventions only. You can filter by region, country, age group, outcome, and location of intervention to identify interventions relevant to your setting."
        ),

        tags$p(
          tags$strong("Studies included: "),
          "Use this tab to explore the full evidence base from 2018 at study level, including interventions that reduced childhood sexual violence, showed no effect or increased childhood sexual violence."
        ),

        modal_heading("Tips for using the dashboard", margin_top = "22px", margin_bottom = "12px", bold = TRUE),

        tags$ul(
          tags$li("Start broad by selecting a region or country, then narrow by age group, outcome, or context."),
          tags$li("Use the View button in tables to open more detailed intervention or study summaries."),
          tags$li("Use the search bar to look up specific studies, interventions, or key words."),
          tags$li("Click countries on the map to see country-level details and download an Excel summary."),
          tags$li("Use the download buttons to export filtered data for reporting or review."),
          tags$li("Use Reset Filters to return to the full dataset.")
        )
      )
    )
  }
  
  # Key definitions modal (used from Home link and Home button)
  definitions_modal <- function() {
    modalDialog(
      title = modal_title("Definitions of terms used in this dashboard"),
      easyClose = TRUE,
      size = "m",
      footer = modalButton("Close"),
      tagList(
        tags$div(
          tags$p(
            tags$strong("Childhood sexual violence (CSV):"),
            "CSV is a preventable violation of ",
            tags$strong("a child\u2019s"),
            " rights, with severe short and long-term consequences for survivors, families, and communities. CSV is defined by UNICEF as any deliberate \u201cact of a sexual nature that is perpetrated against a child [18 years or younger] and that results in or has a high likelihood of resulting in injury, pain, or psychological suffering\u201d (United Nations Children's Fund, 2023)."
          ),
          tags$div(
            tags$small(
              em("Reference: United Nations Children\u2019s Fund. International Classification of Violence against Children. New York: UNICEF, 2023."),
              style = paste0("display:block; margin-top:6px; color:", sfh_charcoal_mid, ";")
            )
          )
        ),
        tags$hr(),
        tags$div(
          tags$p(
            tags$strong("INSPIRE:"),
            "INSPIRE is a set of seven evidence-based strategies for policymakers, practitioners, and researchers with the greatest potential to reduce violence against children. The seven INSPIRE strategies are: Implementation and enforcement of laws; Norms and values; Safe environments; Parent and caregiver support; Income and economic strengthening; Response and support services; Education and life skills."
          ),
          tags$div(
            tags$small(
              em("Reference: World Health Organization. INSPIRE Handbook: action for implementing the seven strategies for ending violence against children. Geneva: World Health Organization, 2018."),
              style = paste0("display:block; margin-top:6px; color:", sfh_charcoal_mid, ";")
            )
          )
        )
      )
    )
  }
  
  # Open the definitions modal from either Home link or Home button
  observeEvent(input$view_definitions_home, { showModal(definitions_modal()) })

  # ---------------------------------------------------------------------------
  # Approach FAQ modal
  # ---------------------------------------------------------------------------
  approach_modal <- function() {
    osf_link <- tags$a(href = "https://osf.io/d2srw/overview", target = "_blank", "here")

    faq_item <- function(question, ...) {
      tags$li(
        style = "margin-bottom:16px;",
        tags$strong(question),
        ...
      )
    }

    faq_answer <- function(...) {
      tags$p(style = "margin:6px 0 0 0;", ...)
    }

    modalDialog(
      title = modal_title("FAQs on our approach"),
      easyClose = TRUE,
      size = "l",
      footer = modalButton("Close"),

      tags$div(
        style = "line-height:1.6; font-family:'General Sans', sans-serif;",

        tags$p(
          "More information on our approach is published ", osf_link, ". Below we share some of the key questions about the approach. We hope this will inform your engagement with the dashboard. Please use the feedback buttons to give us your feedback and the submit evidence button to help us grow this resource."
        ),

        tags$ul(
          style = "padding-left:22px; margin-top:14px;",

          faq_item(
            "Where did the evidence in this dashboard come from?",
            faq_answer("We searched 22 databases for evaluations of interventions, laws, and policies to prevent CSV in any setting from 2018 onwards. The initial search was conducted on 12th October 2025 and yielded 51,728 articles which we screened to identify the studies presented in this dashboard. This is a living review so it will periodically be updated with new evidence.")
          ),

          faq_item(
            "Which types of evaluations are included?",
            faq_answer("All evaluations in this brief measured CSV victimisation and perpetration among children under the age of 18, and perpetrators of CSV at any age. All evaluations were quantitative and published in English in academic peer-reviewed journals.")
          ),

          faq_item(
            "Which types of evidence are missing from this dashboard?",
            faq_answer("This dashboard does not capture the work being done to prevent and respond to CSV that is ongoing, unpublished. Because this is living dashboard we hope to include this work as it is published."),
            faq_answer("The dashboard also does not include qualitative or formative evaluations, prevention efforts where all evaluations were conducted before 2018, and does not provide an assessment of the quality of evidence. This dashboard focuses on sexual violence and does not examine the effects of interventions, laws, or policies on other types of violence against children.")
          ),

          faq_item(
            "How do you decide whether interventions, laws and policies were effective?",
            faq_answer("We examined the authors’ reports of the effectiveness of the intervention. We assessed all CSV victimisation and perpetration outcomes reported in the evaluation and grouped prevention efforts into three groups:"),
            tags$ul(
              style = "list-style-type:circle; padding-left:22px; margin-top:6px; margin-bottom:0;",
              tags$li("Reduced CSV"),
              tags$li("No change in CSV"),
              tags$li("Increased CSV")
            )
          ),

          faq_item(
            "What if an intervention, law, or policy increased some CSV outcomes and reduced others?",
            faq_answer("If an intervention reduced one CSV outcome and increased another CSV outcome, we include it in the ‘Studies included’ tab and do not include it in the ‘Prevention explorer’ tab.")
          ),

          faq_item(
            "What about evidence on other CSV outcomes?",
            faq_answer("We hope to expand this review to include other CSV outcomes. At the moment, the evidence included focuses on whether CSV prevention efforts reduce CSV victimisation and perpetration.")
          ),

          faq_item(
            "What about evidence before 2018?",
            faq_answer("This is a living review, and we hope to expand it to include other years over time. In its current iteration, all interventions in the ‘Prevention explorer’ had evidence from before 2018 checked for any evidence that found an increase to CSV victimisation or perpetration outcomes. Interventions with any published evidence with increased CSV outcomes are not listed in the Prevention explorer tab. This is important because not all the CSV victimisation and perpetration outcomes measured in an evaluation are reported in the same study."),
            faq_answer(
              "If you are interested in evidence before 2018, please have a look at ",
              tags$a(
                href = "https://www.togetherforgirls.org/en/resources/what-works-to-prevent-sexual-violence-against-children-evidence-review",
                target = "_blank",
                "What works to prevent sexual violence against children: Evidence review"
              ),
              "."
            )
          )
        )
      )
    )
  }

  observeEvent(input$open_approach_modal, { showModal(approach_modal()) })
}


# =============================================================================
# 12. App launch
# =============================================================================

shinyApp(ui = ui, server = server)
