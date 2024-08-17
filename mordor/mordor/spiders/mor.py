import scrapy
import pandas as pd
import os

class MorSpider(scrapy.Spider):
    name = "mor"
    allowed_domains = ["mordorintelligence.com"]
    start_urls = ["https://www.mordorintelligence.com/market-analysis/aerospace-defense"]

    def __init__(self, *args, **kwargs):
        super(MorSpider, self).__init__(*args, **kwargs)
        self.data = []

    def parse(self, response):
        for report in response.css('div.single-report-card-wrap'):
            title = report.css('a.report-list-single-card-title::text').get()
            link = report.css('a.report-list-single-card-title::attr(href)').get()

            full_link = response.urljoin(link)
            
            self.data.append({
                'Title': title.strip() if title else '',
                'URL': full_link
            })

        # Handle pagination if there are more pages
        next_page = response.css('a.next::attr(href)').get()
        if next_page:
            yield response.follow(next_page, self.parse)

    def closed(self, reason):
        # Convert data to DataFrame
        df = pd.DataFrame(self.data)

        # File name
        file_name = 'mordorintelligence_reports.xlsx'

        # Check if the Excel file already exists
        if os.path.exists(file_name):
            # Read existing file
            existing_df = pd.read_excel(file_name, engine='openpyxl')
         
            df = pd.concat([existing_df, df], ignore_index=True)

        # Save data to Excel
        df.to_excel(file_name, index=False, engine='openpyxl')
