select distinct
	data.pod_pk,
	pow.pow_pk,
	worksheets.wrk_pk,
	store.sto_name,
	vendor.ven_companyname,
	inventory.INV_Scancode,
	brand.brd_name,
	inventory.INV_ReceiptAlias,
	inventory.inv_size,
	dpt.DPT_Name,
	oi.ord_supplierstocknumber,
	oi.ord_quantityinorderunit,
	data.pod_orderquantity,
	data.pod_shipquantity,
	data.pod_receivequantity,
	pow.pow_invoicenumber, 
	pow.pow_receivedate,
	pow.pow_invoicedate,
	pow.WRK_TimestampCommitted,
	additionalcodes.asc_scancode,
	additionalcodes.asc_quantity,
	worksheets.wrk_name
from momsdatawarehouse.catapult.PurchaseOrderWorksheetData as data  
	inner join momsdatawarehouse.catapult.PurchaseOrderWorksheet as pow on data.pod_pow_fk = pow.pow_pk and data.pod_pow_cfk = pow.pow_cpk
	left join momsdatawarehouse.momcat.store as store on pow.pow_sto_fk_receive = store.sto_pk and pow.pow_sto_cfk_receive = store.sto_cpk
	left join momsdatawarehouse.catapult.Worksheets as worksheets on pow.pow_wrk_fk = worksheets.wrk_pk and pow.pow_wrk_cfk = worksheets.wrk_cpk
	left join momsdatawarehouse.catapult.Vendor as vendor on pow.pow_ven_fk = vendor.ven_pk and pow.pow_ven_cfk = vendor.ven_cpk
	left join momsdatawarehouse.catapult.StockInventory as inventory on data.pod_inv_fk = inventory.inv_pk and data.pod_inv_cfk = inventory.inv_cpk
	left join momsdatawarehouse.catapult.Departments as dpt on inventory.INV_DPT_FK = dpt.dpt_pk and inventory.INV_DPT_CFK = dpt.dpt_cpk
	left join momsdatawarehouse.catapult.Brands as brand on inventory.INV_BRD_FK = brand.brd_pk and inventory.INV_BRD_CFK = brand.brd_cpk
	left join momsdatawarehouse.catapult.OrderingInfo as oi on data.pod_ord_fk = oi.ord_pk and data.pod_ord_cfk = oi.ord_cpk
	left join momsdatawarehouse.catapult.AdditionalScanCodes as additionalcodes on additionalcodes.asc_pk = oi.ord_asc_fk and additionalcodes.asc_cpk = oi.ord_asc_cfk
    left join momsdatawarehouse.catapult.StockInventoryLocal as sil on (CONTAINS_SUBSTR(UPPER(sil.sil_location), "G") or CONTAINS_SUBSTR(UPPER(sil.sil_location), "N"))
where (Date(pow.pow_receivedate) BETWEEN DATE_SUB(current_date()-1, INTERVAL 30 DAY) and current_date()-1) and dpt.dpt_name is not null and worksheets.wrk_committed = True and worksheets.wrk_committed = True and (pow.wrk_timestampcommitted IS NOT NULL and data.wrk_timestampcommitted IS NOT NULL and worksheets.wrk_timestampcommitted IS NOT NULL) and data.pod_receivequantity IS NOT NULL and pow.pow_receivedate IS NOT NULL