# CSV Order Exporter
This script produces a CSV export mimicing the Shopify Admin export. It paginates orders and writes to a CSV file.

## Prerequisites
* Ruby
* Gems: active_support, shopify_api, retriable (`gem install active_support shopify_api retriable`)

## Running the script

The script requires several variables:
* `SHOP` (i.e `myshop` if `https://myshop.myshopify.com` is your shop url)
* `API_KEY` private app api key
* `PASSWORD` private app password
* `DATE` the date to export orders in YYYY-MM-DD format

Example invocation:

`SHOP=myshop API_KEY=key PASSWORD=secret DATE=2018-10-17 ruby export_orders.rb`

the CSV file will be saved in the directory as `myshop_orders_2018-10-17_<timestamp>.csv` where `<timestamp>` is the time the export started.
