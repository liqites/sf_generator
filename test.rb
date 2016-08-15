require 'csv'
require 'roo'
require 'spreadsheet'

# 同时支持CSV和导入

csv_file = Rails.root.to_s + "/script/data/有赞导出订单-月饼-20160812.csv"
excel_file = Rails.root.to_s + "/script/data/顺丰订单#{Time.now.to_i}.xls"

class Item
  attr_accessor :name, :count
end

class Order
  OrderType = {
    0 => 'Single',
    1 => 'Multiple',
    2 => 'Error'
  }
  
  attr_accessor :order_id, :items, :user_name, :user_phone, :user_address, :type, :create_time, :items_count
  
  def initialize
    self.items = []
    self.items_count = 0
    self.type = -1
  end
  
  def items_str
    items.collect{|i| "#{i.name} X #{i.count}"}.join(" & ")
  end
  
  def create_time_str
    create_time.strftime("%F %H:%M")
  end
  
  def to_json
    {
      id: order_id,
      user_name: user_name,
      user_phone: user_phone,
      user_address: user_address,
      items: items_str,
      type: type,
      create_time: create_time_str
    }
  end
end

csv_rows = Roo::CSV.new(csv_file, csv_options: {encoding: Encoding::UTF_16, col_sep: ','}).sheet(0)

order_list = {}
index = 0

# 保存订单列表
csv_rows.each do |cells|
  index += 1
  if index >= 10
    order_id = cells[0][1..-2].strip
    if order_list[order_id].nil?
      user_name = cells[17]
      user_phone = cells[28]
      user_address = cells[21]
      create_time = Time.zone.parse(cells[30] || "") # 如果时间是空的，那么认为这条记录有问题，后面要标出来
      
      order = Order.new
      order.order_id = order_id
      order.user_name = user_name
      order.user_phone = user_phone
      order.user_address = user_address
      order.create_time = create_time
      
      item = Item.new
      item.name = cells[13].strip
      item.count = cells[36].to_i
      order.items << item
      order.items_count += item.count
      
      order_list[order_id] = order
    else
      item = Item.new
      item.name = cells[13].strip
      item.count = cells[36].to_i
      
      order_list[order_id].items << item
      order_list[order_id].items_count += item.count
    end
  end
end

# 处理订单类型
# 通用化这个方法
order_list.values.each do |order|
  if order.items.count > 1
    order.type = 1
  end 

  if order.items.count == 1 && order.create_time.present?
    order.type = 0
  end
  
  if order.type == -1
    order.type = 2
  end
end

o_list = order_list.values

single_lists = {}
o_list.select{ |o| o.type == 0}.each do |o|
  if single_lists[o.items[0].name].nil?
    single_lists[o.items[0].name] = []
  end

  single_lists[o.items[0].name] << o
end

multi_list = o_list.select{ |o| o.type == 1 }
error_list = o_list.select{ |o| o.type == 2 }

# puts small_list.collect(&:to_json)
# puts big_list.collect(&:to_json)
# puts bs_list.collect(&:to_json)

# 对每个订单进行排序
single_lists.each do |key, list|
  list.sort!{|a,b| a.items_count <=> b.items_count}
end

multi_list.sort!{|a,b| a.items_count <=> b.items_count}

# error_list.sort!{|a,b| a.create_time <=> b.create_time}

headers = %w(
 用户订单号
 收件公司 联系人 联系电话	手机号码	收件详细地址
 付款方式 第三方付月结卡号	托寄物内容	托寄物数量	件数	
 实际重量（KG）	计费重量（KG） 业务类型 是否代收货款 代收货款金额
 是否保价 保价金额 标准化包装服务（元） 其它费用（元） 是否自取
 是否签回单	是否定时派送 派送日期 派送时段	是否电子验收
 是否保单配送	是否拍照验证 是否保鲜服务	是否易碎件	是否大闸蟹专递
 是否票据专送	收件员	寄方签名 寄件日期 签收短信通知(MSG)
 派件短信通知(SMS) 寄方客户备注 长(cm) 宽(cm) 高(cm)
 扩展字段1 扩展字段2 扩展字段3
 )


# 创建一个文件
book = Spreadsheet::Workbook.new
sheet1 = book.create_worksheet

format = Spreadsheet::Format.new :size => 14,:weight => :bold

sheet1.row(0).replace headers
sheet1.row(0).default_format = format

format = Spreadsheet::Format.new  :size => 14
                                  
index = 1
(multi_list + single_lists.values.flatten).each do |order|
  sheet1.row(index).replace [
    order.order_id,
    '个人', order.user_name, order.user_phone, '', order.user_address,
    '寄付现结','',order.items_str, order.items_count ,order.items_count,
    '','','顺丰次日','','',
    '','','','','',
    '','','','','',
    '','','Y','','',
    '','','','','',
    'Y','','','','',
    '','',''
                            ]
  sheet1.row(index).default_format = format
  index += 1
end

format = Spreadsheet::Format.new  :size => 14, :pattern_fg_color => :yellow, :pattern => 1
error_list.each do |order|
  sheet1.row(index).replace [
    order.order_id,
    '个人', order.user_name, order.user_phone, '', order.user_address,
    '寄付现结','',order.items_str, order.items_count ,order.items_count,
    '','','顺丰次日','','',
    '','','','','',
    '','','','','',
    '','','Y','','',
    '','','','','',
    'Y','','','','',
    '','',''
                            ]
  sheet1.row(index).default_format = format
  index += 1
end

book.write excel_file