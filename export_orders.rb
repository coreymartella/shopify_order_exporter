require 'shopify_api'
require 'active_support'
require 'active_support/core_ext'
require 'csv'

class CSVOrderExporter
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
  def perform
    raise "SHOP must be specified" unless ENV["SHOP"]
    start = Time.now
    page_size = ENV["PAGE_SIZE"].to_i
    page_size = 250 if page_size < 1 || page_size > 250
    if ENV["TOKEN"]
      session = ShopifyAPI::Session.new("#{ENV["SHOP"]}.myshopify.com", ENV["TOKEN"])
      ShopifyAPI::Base.activate_session(session)
    elsif ENV["API_KEY"] && ENV["PASSWORD"]
      ShopifyAPI::Base.site = "https://#{ENV["API_KEY"]}:#{ENV["PASSWORD"]}@#{ENV["SHOP"]}.myshopify.com/admin"
    else
      raise "TOKEN or API_KEY and PASSWORD must be provided"
    end

    Time.zone = ShopifyAPI::Shop.current.iana_timezone

    date = (Date.parse(ENV["DATE"]) rescue nil)
    raise("Unable to parse date, pass DATE=YYYY-MM-DD") if ENV["DATE"].present? && !date
    date ||= Date.today
    min_time = date.in_time_zone
    max_time = min_time.at_end_of_day

    total = ShopifyAPI::Order.count(params: {created_at_min: min_time.iso8601, created_at_max: max_time.iso8601})
    processed = 0
    page = 1
    more_records = true
    while more_records
      orders = ShopifyAPI::Order.all(params: {page: page, created_at_min: min_time.iso8601, created_at_max: max_time.iso8601, limit: page_size})
      orders.each_with_index do |o,i|
        write_order(o)
        processed += 1
        per_iteration = (Time.now - start)/processed
        est = (start + (per_iteration*total))
        printf("\r%6d/%d (Order %s) ETA: %s (%.2f hours)", processed, total, o.id, est, (est-Time.now)/1.hour);STDOUT.flush
      end
      break if ENV["MAX_RECORDS"] && ENV["MAX_RECORDS"] <= records
      page += 1
    end
    csv.close
    puts "#{Time.now} Wrote #{records} in #{Time.now-start}"
  end
  protected
    def csv
      @csv ||= begin
        csv = CSV.open("#{ENV["SHOP"]}_orders_#{Time.now.strftime("%Y%m%d%H%M%S")}.csv", "wb")
        csv << headers
        csv
      end
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
      transactions = ShopifyAPI::Transaction.find(:all, :params => {order_id: o.id, fields: [:id, :kind, :status, :amount, :created_at, :gateway, :receipt, :source] })
      payment_reference_transaction = transactions.select { |t| %w(authorization sale).include?(t.kind) && t.status = "success" }.last
      payment_reference_receipt = payment_reference_transaction.try(:receipt)
      total_received = transactions.select{|t| %w(capture sale).include?(t.kind) && t.status == "success"}.map(&:amount).map(&:to_d).sum
      total_received -= transactions.select{|t| %w(change).include?(t.kind) && t.status == "success"}.map(&:amount).map(&:to_d).sum
      row = [
       o.name, #
       o.contact_email, # {:label_name=>"Email"}
       o.financial_status || 'pending', #
       transactions.detect{|t| t.kind == "capture" && t.status == "success"}.try(:created_at), # paid_at
       o.fulfillment_status || "unfulfilled", # {:default=>"pending"}
       "", #failure of, #fulfilled_at
       (o.customer.accepts_marketing ? "yes" : "no"), #failure of, #marketing_preference
       o.currency, #
       o.subtotal_price, # {:label_name=>"Subtotal"}
       o.shipping_lines.map(&:price).map(&:to_d).sum, #failure of, #shipping_price
       o.total_tax, # {:label_name=>"Taxes"}
       o.total_price, # {:label_name=>"Total"}
       nil, #failure of, #discount_code
       o.total_discounts, # {:label_name=>"Discount Amount"}
       o.shipping_lines.first&.title, #failure of, #shipping_title
       (o.created_at ? Time.parse(o.created_at).to_s : nil), # {:label_name=>"Created at"}
       o.line_items.first&.quantity,
       o.line_items.first&.name,
       o.line_items.first&.price,
       "", #o.line_items.first&.compare_at_price
       o.line_items.first&.sku,
       o.line_items.first&.requires_shipping,
       o.line_items.first&.taxable,
       o.line_items.first&.fulfillment_status,
       o.billing_address.name.presence || o.customer.name,
       [o.billing_address.address1, o.billing_address.address2].reject(&:blank?).join(", ").presence,
       o.billing_address.address1, # {:label_name=>"Billing Address1"}
       o.billing_address.address2, # {:label_name=>"Billing Address2"}
       o.billing_address.company, # {:label_name=>"Billing Company"}
       o.billing_address.city, # {:label_name=>"Billing City"}
       o.billing_address.zip, # {:label_name=>"Billing Zip", :force_string_in_excel=>true}
       o.billing_address.province_code, # {:label_name=>"Billing Province"}
       o.billing_address.country_code, # {:label_name=>"Billing Country"}
       o.billing_address.phone, # {:label_name=>"Billing Phone"}
       o.shipping_address.name, # {:label_name=>"Shipping Name"}
       [o.shipping_address.address1, o.shipping_address.address2].reject(&:blank?).join(", ").presence,
       o.shipping_address.address1, # {:label_name=>"Shipping Address1"}
       o.shipping_address.address2, # {:label_name=>"Shipping Address2"}
       o.shipping_address.company, # {:label_name=>"Shipping Company"}
       o.shipping_address.city, # {:label_name=>"Shipping City"}
       o.shipping_address.zip, # {:label_name=>"Shipping Zip", :force_string_in_excel=>true}
       o.shipping_address.province_code, # {:label_name=>"Shipping Province"}
       o.shipping_address.country_code, # {:label_name=>"Shipping Country"}
       o.shipping_address.phone, # {:label_name=>"Shipping Phone"}
       o.note, # {:label_name=>"Notes"}
       o.note_attributes.map{ |na| "#{na.name}:, #{na.value}" }.join("\n"),
       (o.cancelled_at ? Time.parse(o.cancelled_at).to_s : nil), # {:label_name=>"Cancelled at"}
       transactions&.first&.gateway, # payment_gateway_for, mismatch of gateway vs provider...
       payment_reference_receipt.try(:trnOrderNumber) || payment_reference_receipt.try(:receipt_id) || payment_reference_transaction.try(:authorization) || payment_reference_transaction.id, #failure of, #payment_reference
       (total_refunded = o.transactions.select{|t| t.kind == "refund" && t.status == "success"}.map(&:amount).map(&:to_d).sum), #failure of, #total_refunded
       o.line_items.first.vendor, # {:label_name=>"Vendor"}
       o.id,
       o.tags,
       "", #failure of, #risk_level_for
       o.source_name, #failure of, #serialized_source_name
       o.line_items.first.total_discount, # {:label_name=>"Lineitem discount"}
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
end

CSVOrderExporter.new.perform
