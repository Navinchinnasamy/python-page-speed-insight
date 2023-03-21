# Import Required Packages
import json
import requests
import pandas as pd
import urllib
import openpyxl
import datetime
# import time

# Get Current Date
cur_date = datetime.date.today()

# Page Speed Insights API Key
API_KEY = "AIzaSyDcVWYyBmSVJLqwWiHFM5rD1nxOAvmXqZU"

# Defining the URL to be analysed
page_url = [
    "https://www.shriramfinance.in",
    "https://www.shriramfinance.in/fixed-deposit",
    "https://www.shriramfinance.in/fixed-deposit-online",
    "https://www.shriramfinance.in/unified-pay-emi",
    "https://www.shriramfinance.in/fixed-deposit-calculator",
    "https://www.shriramfinance.in/press-release",
    "https://www.shriramfinance.in/two-wheeler-loan",
    "https://www.shriramfinance.in/article-which-is-the-best-shriram-loan-scheme-to-buy-a-two-wheeler",
    "https://www.shriramfinance.in/article-gold-loan-vs-credit-card",
    "https://www.shriramfinance.in/article-6-little-known-factors-that-could-affect-your-fixed-deposit-investment",
    "https://www.shriramfinance.in/gold-loan",
    "https://www.shriramfinance.in/business-loan",
    "https://www.shriramfinance.in/recurring-deposits-pay-dues",
    "https://www.shriramfinance.in/downloads",
    "https://www.shriramfinance.in/about-us",
    "https://www.shriramfinance.in/investors/financials",
    "https://www.shriramfinance.in/investors/investor-information",
    "https://www.shriramfinance.in/investors/governance",
    "https://www.shriramfinance.in/investors/fund-raising",
    "https://www.shriramfinance.in/investors/news-announcement",
    "https://www.shriramfinance.in/service-request",
    "https://www.shriramfinance.in/article-how-does-business-loan-work",
    "https://www.shriramfinance.in/article-premature-withdrawal-of-fixed-deposit",
    "https://www.shriramfinance.in/article-weight-of-gst-on-gold-loan-interest",
    "https://www.shriramfinance.in/article-bike-loan-on-fixed-vs-floating-interest-rate-which-is-better",
    "https://www.shriramfinance.in/article-how-different-is-it-to-avail-investment-banking-facilities-from-a-non-banking-firm",
    "https://www.shriramfinance.in/article-are-there-any-extra-charges-for-business-loan-besides-gst-and-interest",
    "https://www.shriramfinance.in/article-which-is-the-best-shriram-finance-loan-scheme-to-buy-a-two-wheeler",
    "https://www.shriramfinance.in/article-gold-loan-eligibility-everything-you-need-to-know",
    "https://www.shriramfinance.in/article-complete-guide-to-get-a-loan-against-an-fd"
]

page_len = len(page_url)

response_object = {}

# DataFrame to store response
df_pagespeed_result = pd.DataFrame(columns=["URL", "Overall_Category", "LCP", "FID", "CLS", "FCP", "Time_to_Interactive", "Total_Blocking_Time", "Speed_Index"])

# Show Progress Bar
def update_progress(progress, length):
    print("\r[{0}] {1}%".format("#" * int((progress)), int((progress/length)*100)))

# Write the API output to a JSON file
def write_to_file(file_name, data):
    with open(file_name, "w+") as fp:
        json.dump(data, fp)

# API Request
def analyse_performance(page_url, strategy):
    response_object[strategy] = {}
    for idx, url_to_check in enumerate(page_url):
        print("## Measuring: {} > {} ##>>".format(idx, url_to_check))
        url = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url="+url_to_check+"&strategy="+strategy+"&key="+API_KEY
        result = urllib.request.urlopen(url).read().decode("UTF-8")
        # Parse response as JSON
        jsonData = json.loads(result)
        response_object[strategy][url_to_check] = jsonData
        print("## Completed: {} > {} ##>>".format(idx, url_to_check))
        if url_to_check != "https://www.shriramfinance.in":
            path_name = url_to_check.replace("https://www.shriramfinance.in/", "").replace("/", "-")
            path_name = path_name+"_{}".format(strategy)
        else:
            path_name = "home_{}".format(strategy)
        json_file_name = "output/{}.json".format(path_name)
        # jsonData = '{ "captchaResult": "CAPTCHA_NOT_NEEDED", "kind": "pagespeedonline#result", "id": "https://www.shriramfinance.in/", "analysisUTCTimestamp": "2023-03-17T07:02:04.556Z"}'
        write_to_file(json_file_name, jsonData)
        update_progress(int(idx+1), page_len)
        # # Sleep for 10 seconds
        # time.sleep(10)

def generate_report(strategy):
    # Generate the DataFrame from the Response Object
    for (url, x) in zip(response_object[strategy].keys(), range(0, len(response_object[strategy]))):
        # URLs
        df_pagespeed_result.loc[x, "URL"] = response_object[strategy][url]["lighthouseResult"]["finalUrl"]

        # Overall Category
        df_pagespeed_result.loc[x, "Overall_Category"] = response_object[strategy][url]["loadingExperience"]["overall_category"]

        ## Core Web Vitals
        # Largest Contentful Paint
        df_pagespeed_result.loc[x, "LCP"] = response_object[strategy][url]["lighthouseResult"]["audits"]["largest-contentful-paint"]["displayValue"]

        # First Input Delay
        if "FIRST_INPUT_DELAY_MS" in response_object[strategy][url]["loadingExperience"]["metrics"]:
            fid = response_object[strategy][url]["loadingExperience"]["metrics"]["FIRST_INPUT_DELAY_MS"]
            df_pagespeed_result.loc[x, "FID"] = fid["percentile"]

        # Cumulative Layout Shift
        df_pagespeed_result.loc[x, "CLS"] = response_object[strategy][url]["lighthouseResult"]["audits"]["cumulative-layout-shift"]["displayValue"]

        # Additional Loading Metrics

        # First Contentful Paint
        df_pagespeed_result.loc[x, "FCP"] = response_object[strategy][url]["lighthouseResult"]["audits"]["first-contentful-paint"]["displayValue"]

        # Time to Interactive
        df_pagespeed_result.loc[x, "Time_to_Interactive"] = response_object[strategy][url]["lighthouseResult"]["audits"]["interactive"]["displayValue"]

        # Total Blocking Time
        df_pagespeed_result.loc[x, "Total_Blocking_Time"] = response_object[strategy][url]["lighthouseResult"]["audits"]["total-blocking-time"]["displayValue"]

        # Speed Index
        df_pagespeed_result.loc[x, "Speed_Index"] = response_object[strategy][url]["lighthouseResult"]["audits"]["speed-index"]["displayValue"]

    # Output
    print(df_pagespeed_result)

    # Generate the Excel Report
    file_name = "report/page_speed_insignts_{}_{}.xlsx".format(cur_date, strategy)
    df_pagespeed_result.to_excel(file_name, sheet_name=str(cur_date))


# Initiate the Page Analysis for Mobile
analyse_performance(page_url, "MOBILE")
generate_report("MOBILE")
analyse_performance(page_url, "DESKTOP")
generate_report("DESKTOP")