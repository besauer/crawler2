from __future__ import annotations

from flask import Flask, render_template, request

from crawler import crawl_site, MAX_PAGES_DEFAULT

app = Flask(__name__)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        raw_start_urls = request.form.get("start_urls", "")
        raw_keywords = request.form.get("keywords", "")
        max_pages = request.form.get("max_pages", type=int, default=MAX_PAGES_DEFAULT)

        start_urls = [url.strip() for url in raw_start_urls.splitlines() if url.strip()]
        keywords = [kw.strip() for kw in raw_keywords.splitlines() if kw.strip()]
        results = []
        error = None

        if not start_urls:
            error = "Bitte geben Sie mindestens eine Start-URL ein."
        elif not keywords:
            error = "Bitte geben Sie mindestens ein Suchstichwort ein."
        else:
            for start_url in start_urls:
                results.extend(crawl_site(start_url, keywords, max_pages=max_pages))

        return render_template(
            "index.html",
            start_urls_input=raw_start_urls,
            keywords_input=raw_keywords,
            max_pages=max_pages,
            results=results,
            error=error,
            submitted=True,
        )

    return render_template(
        "index.html",
        start_urls_input="",
        keywords_input="",
        max_pages=MAX_PAGES_DEFAULT,
        results=[],
        error=None,
        submitted=False,
    )


if __name__ == "__main__":
    app.run(debug=True)
