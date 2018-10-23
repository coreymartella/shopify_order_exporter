# Shopify Canada Order Exporter

This tool exports a spreadsheet (CSV format) of your orders from your
Shopify store. It is a script that runs from the command line tool on
your computer (i.e. Windows Command Prompt, Mac OS Terminal or Linux
Terminal).

## Requirements

To use this software you must:

-   Install the [Ruby programming
    language](https://www.ruby-lang.org/en/documentation/installation/)

-   Install the required Ruby Gems using the command:\
    `gem install activesupport shopify_api retriable`

## Variables

To run the script you will need to provide the following information:

-   `SHOP:` the name of your shop

    -   e.g. If you shop URL is https://myshop.myshopify.com, then your
        shop name is myshop

-   `API_KEY:` Private app API key

    -   [Create this through your Shopify Admin
        page](https://help.shopify.com/en/manual/apps/private-apps).

    -   Use the default values in the **Admin API** and **Storefront API** sections

-   `PASSWORD:` Private app password

    -   This is generated when you create your Private app key

-   `DATE:` The date the orders were placed `YYYY-MM-DD `format

    -   e.g. 2018-10-17

## Running the script

To run the tool:

1.  [Download](https://github.com/coreymartella/shopify_order_exporter/archive/master.zip)
    and open the compressed folder on your computer. The script is in
    this folder.

2.  Open the command line tool on your computer.

3.  Navigate to the folder in the command line.

4.  Type in the following command using the variables described above:

`SHOP=<myshop> API_KEY=<key> PASSWORD=<password> DATE=<date> ruby export_orders.rb`

You should type this all on one line. Do not include the triangle brackets (`< >`).

5.  Press **Enter** or **Return**.

The spreadsheet will be saved in the same folder as the script.
