#!/usr/bin/env ruby

# invoke this script with ENV vars:
# SHOP=myshop API_KEY=private_app_api_key PASSWORD=private_app_password [DATE=YYYY-MM-DD] ruby export_orders.rb

require 'shopify_api'
require 'active_support'
require 'active_support/core_ext'
require 'csv'
require 'retriable'

class CSVOrderExporter
  attr_accessor :date, :since_id
  def initialize
    connect
  end

  def perform
    total = [ShopifyAPI::Order.count(params), max_records].compact.min
    puts "Exporting #{total} orders on #{date} for #{ENV["SHOP"]} to #{filename}"
    processed = 0
    skipped = 0
    page = 1
    more_records = true
    while more_records
      orders = Retriable.retriable{ShopifyAPI::Order.all(params: params.merge(page: page, limit: page_size))}
      orders.each_with_index do |o,i|
        if seen_ids.include?(o.id)
          skipped += 1
          next
        end

        write_order(o)
        seen_ids << o.id

        processed += 1
        per_iteration = (Time.now - start_time)/processed
        est = (start_time + (per_iteration*total))
        printf("\r%6d/%d (Order %14s) ETA: %-20s (%.2f hours)", processed+skipped, total, o.id, est, (est-Time.now)/1.hour);STDOUT.flush

        break if max_records && max_records  <= processed
      end

      break if max_records && max_records  <= processed
      page += 1
      more_records = orders.size == page_size
    end
    csv.close
    puts "#{Time.now} Wrote #{processed} in #{Time.now-start_time} to #{filename}"
  end

  protected
    def date
      @date ||= begin
        d = (Date.parse(ENV["DATE"]) rescue nil)
        raise("Unable to parse date, pass DATE=YYYY-MM-DD") if ENV["DATE"].present? && !d
        d
      end
    end

    def params
      #TODO is created_at_max <= or < ? do we need a +1 on it?
      @params ||= begin
        h = {order: "created_at asc", status: "any"}
        h.merge!(created_at_min: date.in_time_zone.iso8601, created_at_max: date.at_end_of_day.iso8601) if date
        h
      end
    end

    def seen_ids
      @seen_ids ||= Set.new
    end

    def start_time
      @start_time ||= Time.now
    end

    def max_records
      ENV["MAX_RECORDS"].presence&.to_i
    end

    def connect
      raise "SHOP must be specified" unless ENV["SHOP"]
      start = Time.now
      if ENV["TOKEN"]
        session = ShopifyAPI::Session.new("#{ENV["SHOP"]}.myshopify.com", ENV["TOKEN"])
        ShopifyAPI::Base.activate_session(session)
      elsif ENV["API_KEY"] && ENV["PASSWORD"]
        ShopifyAPI::Base.site = "https://#{ENV["API_KEY"]}:#{ENV["PASSWORD"]}@#{ENV["SHOP"]}.myshopify.com/admin"
      else
        raise "TOKEN or API_KEY and PASSWORD must be provided"
      end
      shop = ShopifyAPI::Shop.current
      Time.zone = shop.iana_timezone
      shop
    end

    def filename
      "#{ENV["SHOP"]}_orders_#{date}_#{Time.now.strftime("%Y%m%d%H%M%S")}.csv"
    end

    def csv
      @csv ||= begin
        csv = CSV.open(filename, "wb")
        csv << headers
        csv
      end
    end

    def page_size
      page_size = ENV["PAGE_SIZE"].to_i
      page_size = 50 if page_size < 1 || page_size > 50
      page_size
    end

    def headers
      @headers ||= ATTRIBUTES.map do |field,extras|
        if field == :tax_line_details
          1.upto(5).map{|i| ["Tax #{i} Name","Tax #{i} Value"]}
        else
          extras && extras[:label_name]  ? extras[:label_name] : field.to_s.humanize
        end
      end.flatten
    end

    def write_order(o)
      transactions = Retriable.retriable{ShopifyAPI::Transaction.find(:all, :params => {order_id: o.id, fields: [:id, :kind, :status, :amount, :created_at, :gateway, :receipt, :source] })}

      payment_reference_transaction = transactions.select { |t| %w(authorization sale).include?(t.kind) && t.status = "success" }.last
      payment_reference_receipt = payment_reference_transaction.try(:receipt)
      total_received = transactions.select{|t| %w(capture sale).include?(t.kind) && t.status == "success"}.map(&:amount).map(&:to_d).sum
      total_received -= transactions.select{|t| %w(change).include?(t.kind) && t.status == "success"}.map(&:amount).map(&:to_d).sum
      paid_at = transactions.detect{|t| t.kind == "capture" && t.status == "success"}.try(:created_at)
      row = [
       o.name,
       o.contact_email,
       o.financial_status || 'pending',
       paid_at ? Time.parse(paid_at).to_s : nil,
       o.fulfillment_status || "unfulfilled",
       o.fulfillment_status == "fulfilled" ? o.fulfillments.map{|f| Time.parse(f.created_at)}.max.to_s : "", #GAP fulfilled_at
       (o.try(:customer)&.accepts_marketing ? "yes" : "no"),
       o.currency,
       o.subtotal_price,
       o.shipping_lines.map(&:price).map(&:to_d).sum,
       o.total_tax,
       o.total_price,
       o.discount_applications{|d| d.try(:code)}.compact.first,
       o.total_discounts,
       o.shipping_lines.first&.title,
       (o.created_at ? Time.parse(o.created_at).to_s : nil),
       o.line_items.first&.quantity,
       o.line_items.first&.name,
       o.line_items.first&.price,
       "", #GAP: missing compare_at_price from line item w/o product data
       o.line_items.first&.sku,
       o.line_items.first&.requires_shipping,
       o.line_items.first&.taxable,
       o.line_items.first&.fulfillment_status,
       o.try(:billing_address)&.name.presence || o.try(:customer)&.name,
       [o.try(:billing_address)&.address1, o.try(:billing_address)&.address2].reject(&:blank?).join(", ").presence,
       o.try(:billing_address)&.address1, # {:label_name=>"Billing Address1"}
       o.try(:billing_address)&.address2, # {:label_name=>"Billing Address2"}
       o.try(:billing_address)&.company, # {:label_name=>"Billing Company"}
       o.try(:billing_address)&.city, # {:label_name=>"Billing City"}
       o.try(:billing_address)&.zip, # {:label_name=>"Billing Zip", :force_string_in_excel=>true}
       o.try(:billing_address)&.province_code, # {:label_name=>"Billing Province"}
       o.try(:billing_address)&.country_code, # {:label_name=>"Billing Country"}
       o.try(:billing_address)&.phone,
       o.try(:shipping_address)&.name,
       [o.try(:shipping_address)&.address1, o.try(:shipping_address)&.address2].reject(&:blank?).join(", ").presence,
       o.try(:shipping_address)&.address1,
       o.try(:shipping_address)&.address2,
       o.try(:shipping_address)&.company,
       o.try(:shipping_address)&.city,
       o.try(:shipping_address)&.zip,
       o.try(:shipping_address)&.province_code,
       o.try(:shipping_address)&.country_code,
       o.try(:shipping_address)&.phone,
       o.note,
       o.note_attributes.map{ |na| "#{na.name}:, #{na.value}" }.join("\n"),
       (o.cancelled_at ? Time.parse(o.cancelled_at).to_s : nil),
       # GAP: provider name mismatches gateway for Bambora vs beanstream..
       transactions&.first&.gateway,
       # GAP: payment_reference varies by provider, not sure of the consistent way to get OrderTransaction#name
       payment_reference_receipt.try(:trnOrderNumber) || payment_reference_receipt.try(:receipt_id) || payment_reference_transaction.try(:authorization) || payment_reference_transaction&.id,
       (total_refunded = transactions.select{|t| t.kind == "refund" && t.status == "success"}.map(&:amount).map(&:to_d).sum),
       o.line_items.first.vendor,
       o.id,
       o.tags,
       "", # GAP: need to get OrderRisk for risk_level_for
       o.source_name,
       o.line_items.first.total_discount,
       ] +
      (1..5).map do |tl_i|
        tl = o.tax_lines[tl_i-1]
        #HACK rounded tax rate
        [
          (tl ? "#{tl.title} #{(tl.rate.to_d*100).round}%" : ""),
          tl&.price
        ]
      end.flatten + [o.phone] #
      csv << row
      line_item_data = empty_line_item_data.merge!(
          "Created at" => Time.parse(o.created_at).to_s,
          "Email" => o.contact_email,
          "Name" => o.name,
          "Phone" => o.phone
        )
      o.line_items[1..-1].each do |li|
        LINE_ITEM_FIELDS.each do |field,header|
          line_item_data[header] = li.try(field)
          line_item_data[header] ||= "pending" if field == :fulfillment_status
        end

        csv << line_item_data.values_at(*headers)
      end
    end

    def empty_line_item_data
      (@empty_line_item_data ||= Hash[headers.map {|h| [h,nil]}]).dup
    end

    ATTRIBUTES = [
      [:name],
      [:contact_email, { label_name: 'Email' }],
      [:financial_status],
      [:paid_at, { label_name: 'Paid at' }],
      [:fulfillment_status, { default: 'pending' }],
      [:fulfilled_at, { label_name: 'Fulfilled at' }],
      [:marketing_preference, { label_name: 'Acce pts Marketing' }],
      [:currency],
      [:subtotal_price, { label_name: 'Subtotal' }],
      [:shipping_price, { label_name: 'Shipping' }],
      [:total_tax, { label_name: 'Taxes' }],
      [:total_price, { label_name: 'Total' }],
      [:discount_code],
      [:total_discounts, { label_name: 'Discount Amount' }],
      [:shipping_title, { label_name: 'Shipping Method' }],
      [:created_at, { label_name: 'Created at' }],
      # line items
      ['line_item_quantity', { label_name: 'Lineitem quantity' }],
      ['line_item_name', { label_name: 'Lineitem name' }],
      ['line_item_price', { label_name: 'Lineitem price' }],
      ['line_item_compare_at_price', { label_name: 'Lineitem compare at price' }],
      ['line_item_sku', { label_name: 'Lineitem sku' }],
      ['line_item_requires_shipping', { label_name: 'Lineitem requires shipping' }],
      ['line_item_taxable', { label_name: 'Lineitem taxable' }],
      ['line_item_fulfillment_status', { label_name: "Lineitem fulfillment status" }],
      # billing address
      [:billing_or_customer_name, { label_name: 'Billing Name' }],
      ['billing_address.street', { label_name: 'Billing Street' }],
      ['billing_address.address1', { label_name: 'Billing Address1' }],
      ['billing_address.address2', { label_name: 'Billing Address2' }],
      ['billing_address.company', { label_name: 'Billing Company' }],
      ['billing_address.city', { label_name: 'Billing City' }],
      ['billing_address.zip', { label_name: 'Billing Zip', force_string_in_excel: true }],
      ['billing_address.province_code', { label_name: 'Billing Province' }],
      ['billing_address.country_code', { label_name: 'Billing Country' }],
      ['billing_address.phone', { label_name: 'Billing Phone' }],
      # shipping address
      ['shipping_address.name', { label_name: 'Shipping Name' }],
      ['shipping_address.street', { label_name: 'Shipping Street' }],
      ['shipping_address.address1', { label_name: 'Shipping Address1' }],
      ['shipping_address.address2', { label_name: 'Shipping Address2' }],
      ['shipping_address.company', { label_name: 'Shipping Company' }],
      ['shipping_address.city', { label_name: 'Shipping City' }],
      ['shipping_address.zip', { label_name: 'Shipping Zip', force_string_in_excel: true }],
      ['shipping_address.province_code', { label_name: 'Shipping Province' }],
      ['shipping_address.country_code', { label_name: 'Shipping Country' }],
      ['shipping_address.phone', { label_name: 'Shipping Phone' }],
      [:note, { label_name: 'Notes' }],
      [:extract_note_attributes, { label_name: 'Note Attributes' }],
      [:cancelled_at, { label_name: 'Cancelled at' }],
      [:payment_gateway_for, { label_name: 'Payment Method' }],
      [:payment_reference, { label_name: 'Payment Reference' }],
      [:total_refunded, { label_name: 'Refunded Amount' }],
      # lineitem - vendor
      ['line_items.first.vendor', { label_name: 'Vendor' }],
      [:id],
      [:tags],
      [:risk_level_for, { label_name: 'Risk Level' }],
      [:serialized_source_name, { label_name: 'Source' }],
      # line item discount
      ['line_items.first.total_discount', { label_name: 'Lineitem discount' }],
      # tax types GST/PST etc...
      [:tax_line_details, { calculated_labels: :tax_line_fields }],
      # Additional columns are added to the end to prevent breaking CSV readers that don't rely on column names.
      [:phone]
    ]
    LINE_ITEM_FIELDS = {
      quantity: "Lineitem quantity",
      name: "Lineitem name",
      price: "Lineitem price",
      compare_at_price: "Lineitem compare at price",
      sku: "Lineitem sku",
      requires_shipping: "Lineitem requires shipping",
      taxable: "Lineitem taxable",
      fulfillment_status: "Lineitem fulfillment status",
      vendor: "Vendor",
      total_discount: "Lineitem discount"
    }
end

