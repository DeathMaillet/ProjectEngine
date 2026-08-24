# ProjectEngine GitHub Pages

Static landing page for:

https://tmailletfr.github.io/ProjectEngine/

## Deployment

Recommended deployment method: a dedicated `gh-pages` branch.

1. Create a new orphan branch named `gh-pages`.
2. Copy the contents of this folder to the root of that branch.
3. Commit and push the branch.
4. In the repository, open `Settings → Pages`.
5. Under `Build and deployment`, select `Deploy from a branch`.
6. Select branch `gh-pages` and folder `/ (root)`.
7. Save.

The site is a plain static site. No build process, package manager or dependency is required.

## Included SEO files

- canonical URL
- Open Graph and Twitter metadata
- SoftwareApplication structured data
- robots.txt
- sitemap.xml
- custom 404 page
- favicon
- `.nojekyll`


## v2 corrections

- Hero slogan locked to two lines on desktop.
- Workflow fit panel vertically centered.
- Analytics and S-Curve cards retain their natural image ratios.
- Download CTAs use `download.html`, which resolves the `.xlsm` asset from the latest GitHub release dynamically.


## v3 correction

The stylesheet URL now carries a version query (`styles.css?v=3`) so GitHub Pages and browsers do not keep serving the older cached CSS. Explicit hero colours and stronger desktop/landscape layout rules were also added.

## Privacy-friendly audience measurement

The site uses Cloudflare Web Analytics for aggregate audience and performance measurement.

- no analytics cookies;
- no local storage or persistent visitor identifier;
- no fingerprinting or cross-site tracking;
- no advertising or profiling;
- no consent banner;
- a dedicated `privacy.html` page explains the measurement.

The Cloudflare beacon is installed on `index.html`, `download.html`, `404.html` and `privacy.html`.

