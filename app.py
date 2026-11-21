# Cập nhật bổ sung vào hàm parse_universal_orders trong app.py:

    # 1. Tự động nhận diện sàn kể cả khi TikTok xuất giao diện Tiếng Việt
    if any(k in cols_lower for k in ['mã kiện hàng', 'được shopee trợ giá', 'trạng thái trả hàng/hoàn tiền']):
        platform = 'Shopee'
    elif any(k in cols_lower for k in [
        'sku subtotal after discount', 'tổng phụ sku sau giảm giá',
        'cancelation/return type', 'loại hủy/trả hàng',
        'order substatus', 'trạng thái phụ', 'seller sku', 'sku người bán'
    ]):
        platform = 'TikTok Shop'
    else:
        platform = hint_platform

    # 2. Mở rộng danh sách tên cột cho cả bản Tiếng Việt của TikTok Shop
    col_order = find_col(df, ['Mã đơn hàng', 'Order ID', 'ID đơn hàng'], partial_candidates=['mã đơn', 'order id', 'id đơn'])
    col_status = find_col(df, ['Trạng Thái Đơn Hàng', 'Trạng thái đơn hàng', 'Order Status'], partial_candidates=['trạng thái đơn', 'order status'])
    col_substatus = find_col(df, ['Order Substatus', 'Trạng thái phụ'], partial_candidates=['substatus', 'trạng thái phụ'])
    col_return_type = find_col(df, [
        'Trạng thái Trả hàng/Hoàn tiền', 'Trạng thái trả hàng',
        'Cancelation/Return Type', 'Loại hủy/trả hàng'
    ], partial_candidates=['trả hàng/hoàn tiền', 'return type', 'hủy/trả'])
    col_ret_qty = find_col(df, ['Số lượng sản phẩm được hoàn trả', 'Sku Quantity of return', 'Số lượng SKU hoàn trả'])
    col_prod = find_col(df, ['Tên sản phẩm', 'Tên Sản Phẩm', 'Product Name'], partial_candidates=['tên sản phẩm', 'product name'])
    col_sku = find_col(df, ['Tên phân loại hàng', 'SKU phân loại hàng', 'Variation', 'Biến thể', 'Seller SKU', 'SKU người bán'])
    col_qty = find_col(df, ['Số lượng', 'Quantity'], partial_candidates=['số lượng', 'quantity'])

    col_rev = find_col(
        df,
        exact_case_candidates=[
            'Tổng số tiền Người mua thanh toán',
            'SKU Subtotal After Discount',
            'Tổng phụ SKU sau giảm giá'
        ],
        exact_insensitive_candidates=[
            'Giá ưu đãi',
            'Tổng giá bán (sản phẩm)',
            'SKU Subtotal Before Discount',
            'Tổng phụ SKU trước giảm giá',
            'Order Amount',
            'Tổng số tiền đơn hàng'
        ],
        partial_candidates=['người mua thanh toán', 'sku subtotal after', 'tổng phụ sku', 'giá ưu đãi']
    )
