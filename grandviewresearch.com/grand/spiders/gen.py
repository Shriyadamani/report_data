import scrapy
import pandas as pd
import os

class GenSpider(scrapy.Spider):
    name = "gen"
    allowed_domains = ["grandviewresearch.com"]
    start_urls = ["https://grandviewresearch.com"]

    def start_requests(self):
        # Loop to generate URLs for pagination
        for i in range(1,121):  # Adjust the range as needed
            url = f"https://www.grandviewresearch.com/industry/technology/page:{i}/sort:Report.publish_date/direction:DESC"
            yield scrapy.Request(url=url, callback=self.parse)

    def __init__(self, *args, **kwargs):
        super(GenSpider, self).__init__(*args, **kwargs)
        self.data = []

    def parse(self, response):
        industry = response.css('div.advanced_report_inner h1.iner_about_heading::text').get().strip()
        
        for item in response.css('div.advanced_report_list.full'):
            # Extract report link and title
            report_link = item.css('h3 a::attr(href)').get()
            report_title = item.css('h3 a::text').get()

            # Extract additional details
            publication_date = item.css('span::text').get().strip() if item.css('span::text').get() else 'Not available'

            # Collect the data
            if report_link and report_title:
                full_link = response.urljoin(report_link)
                self.data.append({
                    'Title': report_title,
                    'Date': publication_date,
                    'Industry': industry,
                    'Link': full_link
                })

        # Handle pagination (if there are more pages to follow)
        next_page = response.css('a.next::attr(href)').get()
        if next_page:
            yield response.follow(next_page, self.parse)

    def closed(self, reason):
        # Convert data to DataFrame
        df = pd.DataFrame(self.data)

        # File name
        file_name = 'gvr.xlsx'

        # Check if the Excel file already exists
        if os.path.exists(file_name):
            # Read existing file
            existing_df = pd.read_excel(file_name, engine='openpyxl')
            # Append new data
            df = pd.concat([existing_df, df], ignore_index=True)

        # Save data to Excel
        df.to_excel(file_name, index=False, engine='openpyxl')