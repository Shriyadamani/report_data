import scrapy
import pandas as pd
import os

class MnmSpider(scrapy.Spider):
    name = "mnm"
    allowed_domains = ["marketsandmarkets.com"]

    def start_requests(self):
        # Loop to generate URLs for pagination
        for i in range(0, 28):  # Adjust the range as needed
            url = f"https://www.marketsandmarkets.com/chemicals-market-research-10_{i}.html"
            yield scrapy.Request(url=url, callback=self.parse)

    def __init__(self, *args, **kwargs):
        super(MnmSpider, self).__init__(*args, **kwargs)
        self.data = []

    def parse(self, response):
        # Extract industry from the header
        industry = response.css('div.marketLHS h2.reportName::text').get().strip()
        
        # Extract the report details
        for item in response.css('ul.markeList li'):
            # Extract report link and title
            report_link = item.css('h3.marketLink a::attr(href)').get()
            report_title = item.css('h3.marketLink a::text').get()

            # Extract additional details
            tarck_wrap = item.css('ul.tarckWrap')
            date_published = tarck_wrap.css('li:nth-child(1)::text').get()
            publication_date = date_published.split('Published:')[-1].strip() if date_published else 'Not available'
            price = tarck_wrap.css('li:nth-child(2)::text').get()
            price = price.split('Price:')[-1].strip() if price else 'Not available'

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
        file_name = 'MnM1.xlsx'

        # Check if the Excel file already exists
        if os.path.exists(file_name):
            # Read existing file
            existing_df = pd.read_excel(file_name, engine='openpyxl')
            # Append new data
            df = pd.concat([existing_df, df], ignore_index=True)

        # Save data to Excel
        df.to_excel(file_name, index=False, engine='openpyxl')