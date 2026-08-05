# Tonies Retailer Scraper

Pulls the full US retailer network from the Tonies store locator, then enriches each independent store with a website and contact email.

Two stages, because the locator gives you locations but not the contact data that makes a list usable.

## Stage 1: Scrape locations

```bash
python scrape_tonies.py
```

Walks the store locator and captures name, address, city, state, ZIP, phone, and coordinates for every listed retailer. Output is raw JSON in `output/`.

## Stage 2: Enrich contacts

```bash
python enrich_emails.py output/tonies_raw_<timestamp>.json
```

For each store:

- **Has a website already?** Crawl it for contact email addresses.
- **No website listed?** Search for the business by name and location, resolve the likely official site, then crawl that.

Independent toy and gift shops frequently have no site in the locator, so the search fallback is what turns a partial list into a complete one.

## Output

Enriched JSON and CSV containing the original location fields plus resolved website and any emails found, the format a sales team can import directly.

## Stack

Python · requests · BeautifulSoup

## Notes

Emails collected are business contact addresses published on company websites. Use them in line with CAN-SPAM and applicable local rules.
