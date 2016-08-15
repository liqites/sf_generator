(function () {
  function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
  }

  this.Converter = (function() {  
    function Converter() {}
    
    Converter.order_list;
    Converter.sf_headers = [
      "用户订单号",
      "收件公司", "联系人", "联系电话", "手机号码","收件详细地址",
      "付款方式","第三方付月结卡号","托寄物内容","托寄物数量","件数",	
      "实际重量（KG）", "计费重量（KG）", "业务类型","是否代收货款","代收货款金额",
      "是否保价", "保价金额", "标准化包装服务（元）", "其它费用（元）", "是否自取",
      "是否签回单", "是否定时派送", "派送日期", "派送时段", "是否电子验收",
      "是否保单配送", "是否拍照验证", "是否保鲜服务", "是否易碎件", "是否大闸蟹专递",
      "是否票据专送", "收件员", "寄方签名", "寄件日期", "签收短信通知(MSG)",
      "派件短信通知(SMS)", "寄方客户备注", "长(cm)", "宽(cm)", "高(cm)",
      "扩展字段1", "扩展字段2", "扩展字段3"
    ];
    
    // 转化为顺丰订单
    // data_array = [
    //  [headers],
    //  [values]
    // ]
    Converter.to_sf = function(data_array) {
      index = 0;
      this.order_list = new Object();
      data_array.forEach(function(cells, i, array) {
        index += 1;
        if(index > 1) {
          order_id = cells[0].trim();

          if(Converter.order_list.hasOwnProperty(order_id)) {
            item = new Object();
            item.name = cells[13].trim();
            item.count = parseInt(cells[36]);

            Converter.order_list[order_id].items.push(item)
            Converter.order_list[order_id].items_count += item.count
          }
          else {
            user_name = cells[17];
            user_phone = cells[28];
            user_address = cells[21];
            create_time = new Date(cells[30] || "");
            order = new Object();
            order.order_id = order_id;
            order.user_name = user_name;
            order.user_address = user_address;
            order.create_time = create_time;

            item = new Object();
            item.name = cells[13].trim();
            item.count = parseInt(cells[36]);
            order.items = []
            order.items.push(item)
            order.items_count = item.count
            Converter.order_list[order_id] = order
          }
        }
      });
      
      single_lists = new Object();
      multi_list = []
      error_list = []
      
      for (var key in Converter.order_list) {
        order = Converter.order_list[key]
        if(order.items.length > 1) {
          // Multi
          multi_list.push(order);
        }
        
        if(order.items.length == 1 && order.create_time != undefined ){
          // Single
          if(single_lists.hasOwnProperty(order.items[0].name)) {
            single_lists[order.items[0].name].push(order);
          } else {
            single_lists[order.items[0].name] = [];
            single_lists[order.items[0].name].push(order);
          }
        }
        
        if(order.type == undefined) {
          // Error
          error_list.push(order);
        }
      }
      // console.log(single_lists)
      // console.log(multi_list)
      // console.log(error_list)
      
      full_list = [];
      full_list.push(Converter.sf_headers);

      multi_list = Converter.order_sort(multi_list);
      Converter.orders_sfrows(multi_list, full_list);

      // 对每组订单排序
      for (var k in single_lists) {
        sorted_list = Converter.order_sort(single_lists[k]);
        Converter.orders_sfrows(sorted_list, full_list);
      }
      
      Converter.order_sort(error_list);
      Converter.orders_sfrows(error_list, full_list);
      
      // 转化为订单
      Converter.to_xlsx(full_list);
    };

    Converter.orders_sfrows = function(orders, push_list) {
      orders.forEach(function (order, i, array) {
        push_list.push([
          order.order_id,
          '个人', order.user_name, order.user_phone, '', order.user_address,
          '寄付现结', '', Converter.items_to_str(order.items), order.items_count, order.items_count,
          '', '', '顺丰次晨', '', '',
          '', '', '', '', '',
          '', '', '', '', '',
          '', '', 'Y', '', '',
          '', '', '', '', '',
          'Y', '', '', '', '',
          '', '', ''
        ]);
      });
    }

    Converter.items_to_str = function (items) {
      str_array = [];
      items.forEach(function (item, i, array) {
        str_array.push(item.name + " X " + item.count)
      });
      return str_array.join(" & ");
    }

    Converter.to_xlsx = function (orders) {
      wb = new Workbook();
      ws = Converter.to_worksheet(orders);
      wb.SheetNames.push('AA');
      wb.Sheets['AA'] = ws;
      var wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });
      saveAs(new Blob([Converter.s2ab(wbout)],{type:"application/octet-stream"}), "test.xlsx")
    }

    Converter.s2ab = function (s) {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    }

    Converter.to_worksheet = function (data, opts) {
      var ws = {};
      var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
      for(var R = 0; R != data.length; ++R) {
        for(var C = 0; C != data[R].length; ++C) {
          if(range.s.r > R) range.s.r = R;
          if(range.s.c > C) range.s.c = C;
          if(range.e.r < R) range.e.r = R;
          if(range.e.c < C) range.e.c = C;
          var cell = {v: data[R][C] };
          if(cell.v == null) continue;
          var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
          
          if(typeof cell.v === 'number') cell.t = 'n';
          else if(typeof cell.v === 'boolean') cell.t = 'b';
          else if(cell.v instanceof Date) {
            cell.t = 'n'; cell.z = XLSX.SSF._table[14];
            cell.v = datenum(cell.v);
          }
          else cell.t = 's';
          
          ws[cell_ref] = cell;
        }
      }
      if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
      return ws;
    }

    Converter.order_sort = function (orders) {
      return orders.sort(function (a, b) {
        return a.items_count - b.items_count;
      });
    }

    return Converter;
  })();
}).call(this);
