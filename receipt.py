receipt_template = """
        
                "Money Receipt"

SR No. : 013546
_________________________________________________
Name: Abu Adnan
Date: 06-06-2024
Time: 10.25AM
Required for: Reformer 2
________________________________________________
Item No.   Particulars		Quantity	Remarks
________________________________________________



________________________________________________
						
                        """
print(receipt_template)
# ...

def build_sales_report(sales_data, report_template=receipt_template):
    total_sales = sum(sale["amount"] for sale in sales_data)
    transactions = len(sales_data)
    avg_transaction = total_sales / transactions

    return report_template.format(
        sep=".",
        start_date=sales_data[0]["date"],
        end_date=sales_data[-1]["date"],
        total_sales=total_sales,
        transactions=transactions,
        avg_transaction=avg_transaction,
    )
    return report_template.format(
            sep=".",
            start_date=sales_data[0]["date"],
            end_date=sales_data[-1]["date"],
            total_sales=total_sales,
            transactions=transactions,
            avg_transaction=avg_transaction,
        )