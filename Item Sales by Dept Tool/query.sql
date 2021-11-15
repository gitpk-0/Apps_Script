var SQL_ITEM_SALES = 
"select \
  \
  substr(s.sto_id,0,2) as sto_id,\
  trim(d.DPT_Name) as dept,\
  trim(b.brd_name) as brand,\
  trim(i.inv_receiptalias) as product_name,\
  trim(i.inv_size) as product_size,\
  trim(p1.PI1_Description) as category,\
  trim(p2.PI2_Description) as subcategory_1,\
  trim(p3.PI3_Description) as origin,\
  trim(p4.PI4_Description) as subcategory_2,\
  trim(i.inv_scancode) as upc_plu,\
  k.SIL_Onhand as on_hand,\
  wk.qty,\
  wk.cost,\
  wk.sales,\
  'END_DATE' as range_end,\
  trim(v.VEN_CompanyName) as vendor_name,\
  trim(o.ORD_SupplierStockNumber) as catalog_id,\
  o.ORD_QuantityInOrderUnit as pack_size \
from \
  catapult.Departments d \
  inner join catapult.StockInventory i on d.dpt_pk = i.inv_dpt_fk and d.dpt_cpk = i.inv_dpt_cfk and dpt_number in (DEPARTMENT_NUMBERS) \
  left outer join catapult.StockInventoryLocal k on i.inv_pk = k.SIL_INV_FK and i.inv_cpk = k.SIL_INV_CFK \
  left outer join catapult.Brands b on i.inv_brd_fk = b.brd_pk and i.inv_brd_cfk = b.brd_cpk \
  left outer join catapult.PowerFieldDataInventory1 p1 on i.inv_pi1_fk = p1.PI1_PK and i.inv_pi1_Cfk = p1.PI1_CPK \
  left outer join catapult.PowerFieldDataInventory2 p2 on i.inv_pi2_fk = p2.PI2_PK and i.inv_pi2_Cfk = p2.PI2_CPK \
  left outer join catapult.PowerFieldDataInventory3 p3 on i.inv_pi3_fk = p3.PI3_PK and i.inv_pi3_Cfk = p3.PI3_CPK \
  left outer join catapult.PowerFieldDataInventory4 p4 on i.inv_pi4_fk = p4.PI4_PK and i.inv_pi4_Cfk = p4.PI4_CPK \
  left outer join catapult.Vendor v on k.SIL_VEN_FK_Default = v.VEN_PK and k.SIL_VEN_CFK_Default = v.VEN_CPK \
  left outer join \
  (\
    select ORD_INV_FK, ORD_INV_CFK,ORD_VEN_FK,ORD_VEN_CFK,ORD_SupplierStockNumber,ORD_QuantityInOrderUnit \
    from catapult.OrderingInfo \
    where ORD_Primary = TRUE and ORD_Discontinued = FALSE \
  ) o on i.inv_pk = o.ORD_INV_FK and i.inv_cpk = o.ORD_INV_CFK \
    and k.SIL_VEN_FK_Default = o.ORD_VEN_FK and k.SIL_VEN_CFK_Default = o.ORD_VEN_CFK \
  inner join \
  (\
      select g.isg_sto_fk, g.isg_sto_cfk, t.sit_inv_fk, t.sit_inv_cfk, sum(sit_quantity) as qty, sum(sit_cost) as cost, sum(sit_amount) as sales, date(max(g.isg_endtime)) as max_sales_date \
      from \
        ( \
        select distinct isg_pk,isg_cpk,isg_sto_fk,isg_sto_cfk,isg_endtime \
        from catapult.SummaryItemGroups \
        where isg_endtime between (TIMESTAMP_SUB('START_DATE 00:00:00.000', interval 0 day)) and (TIMESTAMP_SUB('END_DATE 23:59:59.999', interval 0 day)) \
        ) g \
      inner join \
      (\
        select distinct sit_pk,sit_cpk, sit_isg_fk, sit_isg_cfk, sit_inv_fk, sit_inv_cfk, sit_quantity, sit_amount, sit_cost \
        from catapult.SummaryItems \
        where isg_endtime between (TIMESTAMP_SUB('START_DATE 00:00:00.000', interval 0 day)) and (TIMESTAMP_SUB('END_DATE 23:59:59.999', interval 0 day)) \
      ) t on g.isg_pk = t.sit_isg_fk and g.isg_cpk = t.sit_isg_cfk \
      inner join catapult.StockInventory i on t.sit_inv_fk = i.inv_pk and t.sit_inv_cfk = i.inv_cpk \
      inner join catapult.Departments d on i.inv_dpt_fk = d.dpt_pk and i.inv_dpt_cfk = d.dpt_cpk and d.dpt_number IN (DEPARTMENT_NUMBERS) \
      where \
        t.sit_quantity <> 0 \
      group by \
        g.isg_sto_fk, g.isg_sto_cfk, t.sit_inv_fk, t.sit_inv_cfk \
  ) wk on i.inv_pk = wk.sit_inv_fk and i.inv_cpk = wk.sit_inv_cfk and k.SIL_STO_FK = wk.isg_sto_fk and k.SIL_STO_CFK = wk.isg_sto_cfk \
  inner join momcat.store s on k.SIL_STO_FK = s.sto_pk and k.SIL_STO_CFK = s.sto_cpk \
";