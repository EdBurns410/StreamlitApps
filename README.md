# StreamlitApps

![Python](https://img.shields.io/badge/Python-3.x-blue)
![Streamlit](https://img.shields.io/badge/Built%20with-Streamlit-ff4b4b)
![Status](https://img.shields.io/badge/Status-Active%20Playground-success)

A collection of Streamlit apps I use to turn messy real world data into clean, repeatable workflows for marketing, events and sales teams.

Most of these tools started life as “I am not doing this in Excel ever again” moments, then grew into small internal apps that non-technical colleagues can use through a simple web UI.

---

## Contents

* [What these apps are for](#what-these-apps-are-for)
* [Highlights](#highlights)
* [Featured workflows](#featured-workflows)
* [How I use these apps](#how-i-use-these-apps)
* [Tech stack](#tech-stack)
* [Running the apps locally](#running-the-apps-locally)
* [Roadmap](#roadmap)
* [About the creator](#about-the-creator)

---

## What these apps are for

Broadly, the apps in this repo focus on three areas:

1. **Data cleaning and standardisation**

   * Normalising company names and job titles
   * Fixing broken name fields and locations
   * Preparing lead lists so they are consistent across multiple sources

2. **Segmentation and targeting**

   * Splitting large contact lists into useful segments
   * Identifying VIPs, senior decision makers and target accounts
   * Building event-specific segments for outreach and sponsorship

3. **Sponsor and account matching**

   * Matching tech vendors to relevant target accounts
   * Highlighting high value pockets in large datasets
   * Supporting sponsor sales with better, cleaner targeting lists

All of this is wrapped in Streamlit so that the logic lives in Python, while day-to-day users only see an upload button, a few options and a **Run** button.

---

## Highlights

A few examples of what lives in this repo:

* **`app.py`, `appadvanced.py`, `appp.py`**
  General Streamlit front ends where I experiment with layouts and combine multiple workflows into one interface.

* **`Segments.py`, `SegmentsV2.py`, `SegmentBTSNA25Visprom.py`**
  Segmentation tools that turn large delegate or lead lists into structured buckets, for example by region, seniority, company type or event relevance.

* **`SponsorCompanyTargets.py`, `SponsorColdMatch.py`, `SponsorHSMatch.py`, `RackspaceHStargets.py`**
  Matching and targeting helpers for sponsorship and sales, used to line up the right vendors with the right accounts based on rules and filters.

* **`CompanyReference.py`, `fix_names.py`, `LocationSplitter.py`, `best200eachcompany.py`**
  Utility-style scripts that clean and reshape data, for example splitting out locations into structured fields or pulling the “best” contacts per company.

* **`findvp.py`, `findvp_noemail.py`, `findvpother.py`**
  Filters and finders that surface senior contacts (VP and above, or equivalent) from large mixed-seniority datasets.

A lot of these scripts are event or use-case specific, but the underlying patterns are reusable across most outreach or B2B data work.

---

## Featured workflows

To give a clearer sense of how these apps behave in practice, here are some common patterns:

* **VIP delegate segmenter**
  Upload a large attendee or prospect list, map a few key columns (job title, company, region, sector), and generate labelled segments such as `VIP Banking`, `VIP Fintech`, `Strategic Prospects` and `Do Not Target`.

* **Sponsor–account matcher**
  Take a sponsor's ideal customer profile, feed in a broad market list, and return a short list of high-fit accounts with suggested tiers and notes.

* **Lead list normaliser**
  Accept multiple CSV/Excel exports from different systems, align the schema, clean names and locations, de-duplicate, and output a single master file ready for CRM import.

* **AI-assisted classification (early experiments)**
  Use LLMs on top of Streamlit to auto-suggest segments, tag edge cases and explain why a particular contact or company was bucketed in a certain way.

These workflows are designed so non-technical users can repeat them reliably without touching the underlying Python.

---

## How I use these apps

These apps are used to support:

* **Event marketing and VIP delegate programmes**
  Cleaning and segmenting potential attendees, tagging banks and regulated firms, and prioritising outreach.

* **Sponsor and partner targeting**
  Matching sponsors to accounts, building hit lists and turning raw exports into something a salesperson can act on quickly.

* **Day-to-day data automation**
  Taking repetitive spreadsheet work and wrapping it in a simple Streamlit UI so it becomes a one-click internal tool.

Over time I am evolving these into more general templates for data-heavy Streamlit apps, and adding AI-assisted flows on top where it makes sense.

---

## Tech stack

* **Language:** Python 3
* **Framework:** Streamlit
* **Core libraries:**

  * `pandas` for data manipulation
  * `openpyxl` and friends for Excel I/O
  * Standard Python tooling for cleaning, matching and lookups

Some apps also experiment with LLM/AI integrations for classification and smart suggestions on top of the core logic.

---

## Running the apps locally

Most of the apps follow the same pattern.

1. **Clone the repo**

   ```bash
   git clone https://github.com/EdBurns410/StreamlitApps.git
   cd StreamlitApps
   ```

2. **Create and activate a virtual environment** (optional but recommended)

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. **Install base dependencies**

   ```bash
   pip install streamlit pandas
   ```

   Some scripts may require extra packages such as `openpyxl` for Excel support:

   ```bash
   pip install openpyxl
   ```

4. **Run a specific app**

   ```bash
   streamlit run app.py
   ```

   Swap `app.py` for any of the other `.py` files that contain a Streamlit app.

---

## Roadmap

Planned improvements for this repo:

* Group related apps into folders with shared components
* Add example datasets and screenshots for key tools
* Extract common patterns into reusable modules
* Add clearer entry points for non-technical users

This repo is the working playground for my Streamlit and data automation work, and also acts as a public portfolio for the kinds of internal tools I build.

---

## About the creator

I work at the intersection of data, automation and go-to-market operations, building tools that help small teams behave like they have a full data engineering department in the background.

Streamlit is my favourite way to turn those ideas into real, clickable apps that colleagues and clients can use without needing to see the Python underneath.
