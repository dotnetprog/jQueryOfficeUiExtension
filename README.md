# jQueryOfficeUiExtension
a jQuery plugin that extents fabric js components (with this , you can build editable grid from the table component)


## DataSource

The Datasource object is a must to use for using the officeui extensions.

Currently odata is only supported as datasource type

here's some exemples
```javascript
var dt_src = new OfficeUi.dataSource({
              type: 'odata',
              queryOptions: {
                  $filter: 'InvoiceId eq ' + $("#Id").val(),
                  $expand: 'Product'
              },
              url: '/odata/InvoiceLine',
              async: true,
              schema: {
                  key: 'Id'
              },
          });
var products_dt_src = new OfficeUi.dataSource({
    type: 'odata',
    url: '/odata/Products',
    async: false

});
```


documentation on going....


# MAKE SURE YOU USE THE FABRIC JS FROM THE DEPENDENCIES DIRECTORY , SOME FIXES HAS BEEN ADDED TO THIS ONE.
