select products.inv_scancode,
	vendors.ven_name,
	store.sto_name,
	products.department as department1,
	products.category,
	products.subcategory_1,
	products.brand,
	products.description,
	products.inv_receiptalias,
	products.discontinued,
	store.sto_name
from products 
	inner join locations on products.inv_pk = locations.inv_fk and products.inv_cpk = locations.inv_cfk 
	inner join store on locations.sto_fk = store.sto_pk and locations.sto_cfk = store.sto_cpk 
	inner join vendors on locations.ven_fk_default = vendors.ven_pk and locations.ven_cfk_default = vendors.ven_cpk
where store.sto_name IS NOT NULL and ((momcat.locations.loc_name like 'N%' or momcat.locations.loc_name like 'G%') and momcat.locations.sil_discontinued = 0);