(function() {
    function Workbook() {
        if (!(this instanceof Workbook)) return new Workbook();
        this.SheetNames = [];
        this.Sheets = {};
    }

    // polyfill
    if (!String.prototype.includes) {
        String.prototype.includes = function(search, start) {
            'use strict';
            if (typeof start !== 'number') {
                start = 0;
            }

            if (start + search.length > this.length) {
                return false;
            } else {
                return this.indexOf(search, start) !== -1;
            }
        };
    }

    this.Converter = (function() {
        function Converter() {}
        Converter.csv_sku_index = 14;
        Converter.csv_div = Converter.csv_sku_index - 13;
        Converter.yz_headers = ['SKU', '宝贝总数量', '收货人姓名', '联系手机', '收货地址', '订单付款时间', '订单ID/采购单ID']
        Converter.yz_headers_index = {};

        Converter.order_list;
        Converter.sf_headers = [
            "用户订单号",
            "收件公司", "联系人", "联系电话", "手机号码", "收件详细地址",
            "付款方式", "第三方付月结卡号", "托寄物内容", "托寄物数量", "件数",
            "实际重量（KG）", "计费重量（KG）", "业务类型", "是否代收货款", "代收货款金额",
            "是否保价", "保价金额", "标准化包装服务（元）", "其它费用（元）", "是否自取",
            "是否签回单", "是否定时派送", "派送日期", "派送时段", "是否电子验收",
            "是否保单配送", "是否拍照验证", "是否保鲜服务", "是否易碎件", "是否大闸蟹专递",
            "是否票据专送", "收件员", "寄方签名", "寄件日期", "签收短信通知(MSG)",
            "派件短信通知(SMS)", "寄方客户备注", "长(cm)", "宽(cm)", "高(cm)",
            "扩展字段1", "扩展字段2", "扩展字段3"
        ];
        Converter.sf_status = [
            "运输中", "派送失败", "已收件", "未收件"
        ]

        // 有赞订单表投查找器
        Converter.check_yz_headers = function(headers) {
            var need_find = Converter.yz_headers;
            headers.forEach(function(item, index, array) {
                if (need_find.indexOf(item) !== -1) {
                    Converter.yz_headers_index[item] = index;
                    need_find.splice(need_find.indexOf(item), 1)
                }
            });
            if (need_find.length > 0) {
                sweetAlert("文件格式错误", "你的文件中没有包含以下必需的字段: " + need_find.join(","), "error");
            }

            return need_find.length == 0
        }

        // 将Excel妆化为CSV
        Converter.xlsx_to_csv = function(file, callback) {
            var reader = new FileReader();
            var name = file.name;
            reader.onload = function(e) {
                var data = e.target.result;
                var workbook = XLSX.read(data, { type: 'binary' });

                var sheet_name_list = workbook.SheetNames;
                sheet_name_list.forEach(function(y) { /* iterate through sheets */
                    var worksheet = workbook.Sheets[y];
                    // 转化为csv string
                    var csv_string = XLSX.utils.sheet_to_csv(worksheet);
                    Papa.parse(csv_string, {
                        encoding: 'gb18030',
                        complete: function(obj) {
                            if (callback) {
                                callback(obj.data);
                            }
                            // array = obj.data;
                            // Converter.sf_to_yz_csv(obj.data)
                            // if (callback != undefined) {
                            //   callback();
                            // }
                        },
                        error: function(error) {
                            console.log(error);
                        }
                    });
                });
            };
            reader.readAsBinaryString(file);
        }

        // 读取顺丰文件
        Converter.read_sf_file = function(file, callback) {
            ext = file.name.split('.').pop();
            if (ext == "xlsx" || ext == "xls") {
                console.log("xlsx/xls");
                var reader = new FileReader();
                var name = file.name;
                reader.onload = function(e) {
                    var data = e.target.result;
                    var workbook = XLSX.read(data, { type: 'binary' });

                    var sheet_name_list = workbook.SheetNames;
                    sheet_name_list.forEach(function(y) { /* iterate through sheets */
                        var worksheet = workbook.Sheets[y];
                        // 转化为csv string
                        var csv_string = XLSX.utils.sheet_to_csv(worksheet);
                        Papa.parse(csv_string, {
                            encoding: 'gb18030',
                            complete: function(obj) {
                                Converter.sf_to_yz_csv(obj.data)
                                if (callback != undefined) {
                                    callback();
                                }
                            },
                            error: function(error) {
                                console.log(error);
                            }
                        });
                    });
                };
                reader.readAsBinaryString(file);
            }
        };

        // 顺丰转有赞
        Converter.sf_to_yz_csv = function(data_array) {
            csv_string = ""
            csv_string = csv_string.concat("订单ID,物流公司,物流单号\n");
            index = 0;
            data_array.forEach(function(cells, i, array) {
                index += 1;
                if (index > 1 && cells.length > 10) {
                    csv_string = csv_string.concat(cells[0] + ",顺丰速运," + cells[1] + "\n");
                }
            });
            // console.log(csv_string);
            var blob = new Blob([csv_string], { type: "text/csv;charset=utf-8" });
            saveAs(blob, "标记发货-" + (new Date()).getTime() + ".csv");
        }

        // 读取有赞文件
        Converter.read_yz_file = function(file, callback) {
            var ext = file.name.split('.').pop();
            if (ext == "xlsx" || ext == "xls") {
                console.log("xlsx/xls");
                var reader = new FileReader();
                var name = file.name;
                reader.onload = function(e) {
                    var data = e.target.result;
                    var workbook = XLSX.read(data, { type: 'binary' });

                    var sheet_name_list = workbook.SheetNames;
                    sheet_name_list.forEach(function(y) { /* iterate through sheets */
                        var worksheet = workbook.Sheets[y];
                        // 转化为csv string
                        var csv_string = XLSX.utils.sheet_to_csv(worksheet);
                        Papa.parse(csv_string, {
                            encoding: 'gb18030',
                            complete: function(obj) {
                                Converter.csv_to_sf(obj.data)
                                if (callback != undefined) {
                                    callback();
                                }
                            },
                            error: function(error) {
                                console.log(error);
                            }
                        });
                    });
                };
                reader.readAsBinaryString(file);
            }

            if (ext == "csv") {
                console.log("CSV");
                Papa.parse(file, {
                    encoding: 'gb18030',
                    complete: function(obj) {
                        Converter.csv_to_sf(obj.data)
                        if (callback != undefined) {
                            callback();
                        }
                    },
                    error: function(error) {
                        console.log(error);
                    }
                });
            }
        }

        // 有赞CSV转化为顺丰
        // ['SKU', '宝贝总数量', '收货人姓名', '联系手机', '收货地址', '订单付款时间', '订单ID/采购单ID']
        Converter.csv_to_sf = function(data_array) {
            var order_id_idx, item_name_idx, item_count_idx, user_name_idx, user_address_idx, user_phone_idx, order_create_time_idx;
            var index = 0;
            this.order_list = new Object();
            data_array.forEach(function(cells, i, array) {
                index += 1;
                // console.log(cells);

                // 检查列名
                if (index == 1) {
                    if (Converter.check_yz_headers(cells)) {
                        order_id_idx = Converter.yz_headers_index['订单ID/采购单ID'],
                            item_name_idx = Converter.yz_headers_index['SKU'],
                            item_count_idx = Converter.yz_headers_index['宝贝总数量'],
                            user_name_idx = Converter.yz_headers_index['收货人姓名'],
                            user_phone_idx = Converter.yz_headers_index['联系手机'],
                            user_address_idx = Converter.yz_headers_index['收货地址'],
                            order_create_time_idx = Converter.yz_headers_index['订单付款时间'];
                    } else {

                    }
                    // if (cells[Converter.csv_sku_index].trim() != 'SKU') {
                    //     sweetAlert("文件格式错误", "你选择的文件第" + Converter.csv_sku_index + 1 + "列不是SKU，会导致转化出错，请检查并修改为正确的格式后重新选择！", "error");
                    //     // return;
                    // }
                }



                if (index > 1 && cells.length > 10) {
                    order_id = cells[order_id_idx].trim();

                    if (Converter.order_list.hasOwnProperty(order_id)) {
                        item = new Object();
                        item.name = cells[item_name_idx].trim();
                        item.count = parseInt(cells[item_count_idx]);

                        Converter.order_list[order_id].items.push(item)
                        Converter.order_list[order_id].items_count += item.count
                    } else {
                        user_name = cells[user_name_idx];
                        user_phone = cells[user_phone_idx];
                        user_address = cells[user_address_idx];
                        create_time = new Date(order_create_time_idx);
                        order = new Object();
                        order.order_id = order_id;
                        order.user_name = user_name;
                        order.user_phone = user_phone;
                        order.user_address = user_address;
                        order.create_time = create_time;

                        item = new Object();
                        item.name = (cells[item_name_idx] || "").trim();
                        item.count = parseInt(item_count_idx);
                        order.items = []
                        order.items.push(item)
                        order.items_count = item.count
                        Converter.order_list[order_id] = order
                    }
                }
            });

            var single_lists = new Object();
            var multi_list = [];
            var error_list = [];

            for (var key in Converter.order_list) {
                var order = Converter.order_list[key]
                if (order.items.length > 1) {
                    // Multi
                    multi_list.push(order);
                } else if (order.items.length == 1 && order.create_time != undefined) {
                    // Single
                    if (single_lists.hasOwnProperty(order.items[0].name)) {
                        single_lists[order.items[0].name].push(order);
                    } else {
                        single_lists[order.items[0].name] = [];
                        single_lists[order.items[0].name].push(order);
                    }
                } else {
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
            Converter.save_to_xlsx(full_list);
        };

        Converter.orders_sfrows = function(orders, push_list) {
            orders.forEach(function(order, i, array) {
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

        Converter.items_to_str = function(items) {
            str_array = [];
            items.forEach(function(item, i, array) {
                str_array.push(item.name + " X " + item.count)
            });
            return str_array.join(" & ");
        }

        Converter.save_to_xlsx = function(orders) {
            wb = new Workbook();
            ws = Converter.to_worksheet(orders);
            wb.SheetNames.push('SF');
            wb.Sheets['SF'] = ws;
            var wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });
            saveAs(new Blob([Converter.s2ab(wbout)], { type: "application/octet-stream" }), "顺丰单" + (new Date()).getTime() + ".xlsx");
        }

        Converter.s2ab = function(s) {
            var buf = new ArrayBuffer(s.length);
            var view = new Uint8Array(buf);
            for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        }

        Converter.to_worksheet = function(data, opts) {
            var ws = {};
            var range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
            for (var R = 0; R != data.length; ++R) {
                for (var C = 0; C != data[R].length; ++C) {
                    if (range.s.r > R) range.s.r = R;
                    if (range.s.c > C) range.s.c = C;
                    if (range.e.r < R) range.e.r = R;
                    if (range.e.c < C) range.e.c = C;
                    var cell = { v: data[R][C] };
                    if (cell.v == null) continue;
                    var cell_ref = XLSX.utils.encode_cell({ c: C, r: R });

                    if (typeof cell.v === 'number') cell.t = 'n';
                    else if (typeof cell.v === 'boolean') cell.t = 'b';
                    else if (cell.v instanceof Date) {
                        cell.t = 'n';
                        cell.z = XLSX.SSF._table[14];
                        cell.v = datenum(cell.v);
                    } else cell.t = 's';

                    ws[cell_ref] = cell;
                }
            }
            if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
            return ws;
        }

        Converter.order_sort = function(orders) {
            return orders.sort(function(a, b) {
                return a.items_count - b.items_count;
            });
        }

        return Converter;
    })();
}).call(this);