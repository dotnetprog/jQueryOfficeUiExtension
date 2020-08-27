
///<reference path="../lib/jquery/dist/jquery.js" />

(function (OfficeUi) {
    OfficeUi.dataSource = function (options) {
        this.Data = [];
        this.schema = options.schema;
        this.defaultorderbyQueryOption = !!options.queryOptions && !!options.queryOptions.$orderby ? options.queryOptions.$orderby : null
        var that = this;
        var asynchronous = options.async === undefined || options.asyncs === null ? true : options.async;
        this.buildNewRowData = function (r) {
            if (!r) {
                r = {}; 
            }
            r.set = function (fn, fv) {
                r[fn] = fv;
                !!that.onValueChanged && that.onValueChanged(r, fn, fv);
            };
            return r;
        };
        this.fetch = function (callback,sort,paging) {
            if (!!this.displayLoading)
                this.displayLoading();
            switch (options.type) {
                case "odata":
                    var u = options.url;
                    if (!!sort) {
                        options.queryOptions = options.queryOptions || {};
                        options.queryOptions.$orderby = sort.field + " " + sort.type;
                    } else {
                        if (!!that.defaultorderbyQueryOption)
                            options.queryOptions.$orderby = that.defaultorderbyQueryOption;
                        else {
                            !!options.queryOptions && !!options.queryOptions.$orderby && delete options.queryOptions.$orderby;
                        }
                    } 

                    if (!!paging) {
                        paging.currentPage = paging.currentPage || 1;
                        options.queryOptions = options.queryOptions || {};
                        options.queryOptions.$top = paging.size;
                        if (paging.currentPage > 1) {
                            options.queryOptions.$skip = paging.size * (paging.currentPage - 1);
                        } else {
                            !!options.queryOptions.$skip && delete options.queryOptions.$skip;
                        }
                        
                    }

                    if (!!options.queryOptions) {
                        var idx = 0;
                        for (var qo in options.queryOptions) {
                            var del = idx === 0 ? '?' : '&';
                            u = u + del + qo + '=' + options.queryOptions[qo];
                            idx++;
                        }

                    }
                   
                   

                    $.ajax({
                        type: 'GET',
                        url: u,
                        dataType: 'json',
                        async: asynchronous,
                        contentType: "application/json",
                        success: function (resp) {


                            that.Data = resp.value.map(
                                function (d) {
                                    d.set = function (fn, fv) {
                                        d[fn] = fv;
                                        !!that.onValueChanged && that.onValueChanged(d, fn, fv);
                                    }
                                    return d;
                                }
                            );
                            callback(resp.value);
                            if (!!that.hideLoading)
                                that.hideLoading();
                        },
                        error: function (err) {
                            console.error(err);
                        }
                    });

                    break;
                default:
                    this.Data = options.loadData();
                    callback(this.Data);
                    if (!!this.hideLoading)
                        this.hideLoading();
                    break;
            }




        };
        const internalProps = ['mode','uid','odata.context'];
        this.create = !!options.create ? options.create : function (datarow,isLast,cb) {
            switch (options.type) {
                case 'odata':
                    var r = {};

                    for (var prop in datarow) {
                        if (internalProps.indexOf(prop) > -1)
                            continue;
                        r[prop] = datarow[prop];
                    }
                    var u = options.url;
                    $.ajax({
                        type: 'POST',
                        url: u,
                        dataType: 'json',
                        data: JSON.stringify(r),
                        async: asynchronous,
                        contentType: "application/json",
                        success: function (resp) {
                            console.log(resp);
                            for (var prop in resp) {
                                if (internalProps.indexOf(prop) > -1)
                                    continue;

                                datarow[prop] = resp[prop];
                            }
                            datarow.mode = null;
                            //TODO Update datarow
                            if(!!cb)
                                cb(datarow, 'new', isLast);
                           
                        },
                        error: function (err) {
                            console.error(err);
                        }
                    });
                    break;
                default:
                    break;
            }
           
            

        };
        this.delete = !!options.delete ? options.delete : function (datarow, isLast ,cb) {
            switch (options.type) {
                case 'odata':
                   
                    var u = options.url + '(' + datarow[options.schema.key] + ')';
                    $.ajax({
                        type: 'DELETE',
                        url: u,
                        dataType: 'json',
                        async: asynchronous,
                        contentType: "application/json",
                        success: function (resp) {

                            var idx = that.Data.indexOf(datarow);
                            that.Data.splice(idx, 1);
                            if(!!cb)
                                cb(datarow, 'removed', isLast);
                        },
                        error: function (err) {
                            console.error(err.responseText);
                            alert(err.responseText);
                        }
                    });
                    break;
                default:
                    break;
            }
        };
        this.update = !!options.update ? options.update :function (datarow, isLast,cb) {
            switch (options.type) {
                case 'odata':
                    var r = {};

                    for (var prop in datarow) {
                        if (internalProps.indexOf(prop) > -1)
                            continue;
                        r[prop] = datarow[prop];
                    }
                    var u = options.url + '(' + datarow[options.schema.key] + ')';
                    $.ajax({
                        type: 'PUT',
                        url: u,
                        dataType: 'json',
                        data: JSON.stringify(r),
                        async: asynchronous,
                        contentType: "application/json",
                        success: function (resp) {
                            console.log(resp);
                            for (var prop in resp) {
                                if (internalProps.indexOf(prop) > -1)
                                    continue;

                                datarow[prop] = resp[prop];
                            }
                            datarow.mode = null;
                            //TODO Update datarow
                            if (!!cb)
                                cb(datarow, 'edited', isLast);
                        },
                        error: function (err) {
                            console.error(err.responseText);
                            alert(err.responseText);
                        }
                    });
                    break;
                default:
                    break;
            }
        };
        this.saveChanges = function (cb) {
            var dirtyRows = that.Data.filter(function (d) { return !!d.mode; });
            for (var i = 0; i < dirtyRows.length; i++) {
                var dr = dirtyRows[i];
                var isLast = i === dirtyRows.length - 1;
                switch (dr.mode) {
                    case 'new':
                        this.create(dr, isLast,cb);

                        break;
                    case 'edited':
                        this.update(dr, isLast,cb);
                        break;
                    case 'removed':
                        this.delete(dr, isLast,cb);
                        break;
                }


            }
        };
        this.getDirtyRows = function () {
            var dirtyRows = that.Data.filter(function (d) { return !!d.mode; });
            return dirtyRows;
        };
        this.getCount = !!options.getCount ? options.getCount : function (cb) {
            switch (options.type) {
                case 'odata':
                    if (!options.odata || !options.odata.counturl) {
                        cb(that.Data.length);
                        return;
                    }
                        
                    var u = options.odata.counturl;
                    //Invoices(87d82f34-3c5b-4405-b173-bdb0e6c4c252)/InvoiceLine/$count
                    
                    $.ajax({
                        type: 'GET',
                        url: u,
                        dataType: 'json',
                        async: asynchronous,
                        contentType: "application/json",
                        success: function (resp) {
                            
                            cb(resp);
                           
                        },
                        error: function (err) {
                            console.error(err);
                        }
                    });
                    break;
                default:
                    cb(that.Data.length);
            }
        };

        
        
    };
   
}(window.OfficeUi = window.OfficeUi || {}));

(function ($) {

    $.fn.OfficeEditableTable = function (config) {
        function uuid() {
            return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
                var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
                return v.toString(16);
            });
        }
        var sortIcons = {
            asc: 'ms-Icon ms-Icon--SortUp',
            desc: 'ms-Icon ms-Icon--SortDown',
            default: 'ms-Icon ms-Icon--Sort'
        };
        var that = this;
        var datasource = config.datasource;
        var columns = config.columns;
        var commandbar = config.commandbar;
        var selectable = config.selectable;
        var editable = !config.IsReadOnly;
        var events = config.events;
        config.schema = datasource.schema;
        var paging = config.paging;
        var schema = config.schema;
        var rowHeight = 46;
        var table_uid = uuid();
        const showDirty = function () {

            if (datasource.getDirtyRows().length > 0)
                $('#' + table_uid + '_save > span.ms-CommandButton-label').text('Unsaved changes');
            else {
                $('#' + table_uid + '_save > span.ms-CommandButton-label').text('Saved');
            }
        };
        const showSaving = function (isSaved) {
            if (!isSaved) {
                $('#' + table_uid + '_save').attr('disabled');
                $('#' + table_uid + '_save > span.ms-CommandButton-label').text('Saving...');
            } else {
                $('#' + table_uid + '_save').removeAttr('disabled');
                $('#' + table_uid + '_save > span.ms-CommandButton-label').text('Saved');
            }

        };

        datasource.onValueChanged = function (d, fn, fv) {
            showDirty();
        }
        const rowFactory = function (recid) {
            var rec = {};
            if (!!recid) {
                rec = datasource.Data.filter(function (d) {
                    return d[schema.key] === recid;
                })[0];
            } else {
                
                rec = {
                    mode: 'new',
                    uid: uuid()
                };
                rec = datasource.buildNewRowData(rec);
                for (var i = 0; i < columns.length; i++) {
                    var c = columns[i];
                    if (!!c.defaultvalue && !!c.field) {
                        rec[c.field] = c.defaultvalue;

                    }
                }

                datasource.Data.push(rec);
            }
            return rec;

        }
        const buildeditableCell = function (c, data) {
            var cell = document.createElement('td');
            if (!!c.field)
                cell.setAttribute('data-attr', c.field);
            else if (!!c.label && !!c.calculated) {
                cell.setAttribute('data-attr', c.label);
            }
            if (c.hidden !== true) {
                if (!!c.editor && typeof (c.editor) === 'object') {
                    if (!!c.editor.template)
                        c.editor.template(cell, data);

                } else {
                    switch (c.type) {
                        default:

                            var input_container = document.createElement('div');
                            input_container.className = "ms-TextField" + (!!c.calculated || c.IsReadOnly ? " is-disabled" : "");
                            var input = document.createElement('input');
                            input.type = c.type == 'decimal' ? 'number' : 'text';
                            input.className = 'ms-TextField-field';
                            input.readOnly = !!c.calculated || c.IsReadOnly;

                            if (!!data[c.field])
                                input.value = data[c.field];
                            else if (!!c.calculated)
                                input.value = c.calculated(data).toString();
                            input_container.append(input);
                            if (!c.calculated && !c.IsReadOnly) {
                                input.addEventListener('change', function (e) {
                                    var ctr = $(e.currentTarget).closest('tr');
                                    var uid = ctr.attr('uid');

                                    var parentCell = $(e.currentTarget).closest('td');
                                    var field = parentCell.attr('data-attr');

                                    var dr = datasource.Data.filter(function (d) { return d.uid === uid; })[0];
                                    var cc = config.columns.filter(function (co) { return co.field === field; })[0];
                                    var nv = $(e.currentTarget).val();
                                    switch (cc.type) {
                                        case 'decimal':
                                            dr[field] = !!nv ? parseFloat($(e.currentTarget).val()) : null;
                                            break;
                                        default:
                                            dr[field] = $(e.currentTarget).val() || "";
                                            break;

                                    }
                                    showDirty();
                                    var columnsToRerender = config.columns.filter(function (co) {
                                        if (!co.calculated)
                                            return false;

                                        return (!!co.triggerFields && co.triggerFields.indexOf(field) > -1);


                                    });
                                    for (var i = 0; i < columnsToRerender.length; i++) {
                                        var cr = columnsToRerender[i];
                                        var td = ctr.children('td[data-attr="' + cr.label + '"]');
                                        var ncell = buildeditableCell(cr, dr);
                                        td.empty().append(ncell.children);
                                        //td.html(ncell.innerHTML);
                                    }


                                });
                            }

                            cell.append(input_container);
                            break;
                    }
                }

            } else {
                cell.style.display = 'none';
                if (!!data[c.field])
                    cell.innerHtml = data[c.field];

            }
            return cell;


        };
        const buildEditableRow = function (dataRow) {
            var ntr = document.createElement('tr');

            if (selectable) {
                var td_select = document.createElement("td");
                td_select.className = "ms-Table-rowCheck";
                ntr.append(td_select);
            }


            ntr.setAttribute('uid', dataRow.uid);
            if (dataRow.mode === 'new')
                ntr.setAttribute('mode', 'new');


            for (var i = 0; i < columns.length; i++) {
                var c = columns[i];
                var cell = buildeditableCell(c, dataRow);
                ntr.append(cell);


            }
            return ntr;
        };
        const setResizeListeners = function(div) {
            var pageX, curCol, nxtCol, curColWidth, nxtColWidth;
           // var table = document.getElementById(table_uid);
            div.addEventListener('mousedown', function (e) {
                curCol = e.target.parentElement;
                nxtCol = curCol.nextElementSibling;
                pageX = e.pageX;
                curColWidth = curCol.offsetWidth
                if (nxtCol)
                    nxtColWidth = nxtCol.offsetWidth
            });

            document.addEventListener('mousemove', function (e) {
                if (curCol) {
                    var diffX = e.pageX - pageX;

                    if (nxtCol)
                        nxtCol.style.width = (nxtColWidth - (diffX)) + 'px';

                    curCol.style.width = (curColWidth + diffX) + 'px';
                }
            });

            document.addEventListener('mouseup', function (e) {
                curCol = undefined;
                nxtCol = undefined;
                pageX = undefined;
                nxtColWidth = undefined;
                curColWidth = undefined;
            });
        }
        const removeAllSortState = function (but) {
            var columns = $('#' + table_uid + ' > thead > tr > th').not(but);
            columns.removeAttr('sortstate');
            columns.children('i').removeClass().addClass(sortIcons.default);
        };
        const buildHeader = function() {
            var header = document.createElement('thead');
            var trHead = document.createElement('tr');
            header.className = "officeuiExtension";
            if (selectable) {
                var th_select = document.createElement("th");
                th_select.className = "ms-Table-rowCheck";
                trHead.append(th_select);
                th_select.addEventListener('click', function (e) {
                    var classAttr = $(this).closest('tr').attr('class');
                    var classList = [];
                    if (!!classAttr) {
                        classList = classAttr.split(/\s+/);
                    }
                    
                    var isSelected = classList.indexOf('is-selected') > -1;
                    var trs = $('#' + table_uid + ' > tbody > tr');
                    if (!isSelected) {
                        trs.addClass('is-selected');
                    } else {
                        trs.removeClass('is-selected');
                    }
                });
            }
            for (var i = 0; i < columns.length; i++) {
                var c = columns[i];
                var th = document.createElement('th');
                th.class = c.class || "";
                if (c.hidden === true)
                    th.style.display = 'none';
                if (!!c.width)
                    th.style.width = c.width;
                


                th.innerText = c.label;
                if (c.sortable && c.sort) {
                    th.className = "officesortable";
                    th.addEventListener('click', function (e) {
                        var clickedelement = e.target;
                        if (clickedelement.className === "colresize")
                            return;
                        if (datasource.getDirtyRows().length > 0) {
                            alert('Save changes before change page');
                            return;
                        }

                        var th = $(this);
                        var sortState = th.attr('sortstate');
                        var childIcon = th.children('i');
                        if (!sortState || sortState == "default") {
                            th.attr('sortstate', 'desc');
                            childIcon.removeClass(sortIcons.default.split(' ')).addClass(sortIcons.desc.split(' '));

                        } else if (sortState === 'desc') {
                            th.attr('sortstate', 'asc');
                            childIcon.removeClass(sortIcons.desc.split(' ')).addClass(sortIcons.asc.split(' '));
                        } else {
                            th.removeAttr('sortstate');
                            childIcon.removeClass(sortIcons.asc.split(' ')).addClass(sortIcons.default.split(' '));
                        }
                        removeAllSortState(th);
                        //add datasource sort logic here.
                        refreshData();

                    });
                    th.setAttribute('sortfield', c.sort.sortedfield || c.field);
                    var iconEl = document.createElement('i');
                    iconEl.className = "ms-Icon ms-Icon--Sort";
                    iconEl.style.float = "right";
                    th.append(iconEl);
                }
                if (c.resizable) {
                    var resizeDiv = document.createElement('div');
                    resizeDiv.className = "colresize";
                    setResizeListeners(resizeDiv);
                    th.append(resizeDiv);
                }
                trHead.append(th);
            }
            header.append(trHead);
            return header;
        }

        const buildTableRow = function (record) {
            if(!record.uid)
                record.uid = uuid();
            var tr = document.createElement("tr");
            tr.setAttribute('uid', record.uid);
            tr.setAttribute('dataid', record[schema.key]);
            if (editable || (!!events && events.onRowDblClick)) {
                tr.addEventListener('dblclick', function (e) {
                    var utr = $(this);
                    var uid = utr.attr('uid');
                    var dr = datasource.Data.filter(function (datarow) { return datarow.uid === uid; })[0];
                    if (!!events && !!events.onRowDblClick) {
                        events.onRowDblClick(dr);
                        return;
                    }
                    
                    if (utr.attr('mode') === 'edit')
                        return;

                    utr.attr('mode', 'edited');

                   

                   
                    dr.mode = 'edited';


                    var new_utr = buildEditableRow(dr);
                    utr.replaceWith(new_utr);



                });
            }
            if (selectable === true) {
                var td_selectable = document.createElement('td');
                td_selectable.className = "ms-Table-rowCheck";
                tr.append(td_selectable);
            }
            for (var i = 0; i < columns.length; i++) {
                var column = columns[i];
                var td = document.createElement('td');
                td.id = record.uid + "_" + column.field;
                td.className = column.cellClass || "";
                if (!!column.template)
                    td.innerHTML = column.template(record);
                else {
                    var v = !!column.calculated ? column.calculated(record) : (record[column.field] || "");
                    switch (column.type) {
                        case 'datetime':
                            var d = new Date(v);
                            if (!!v && !!column.format && typeof (column.format) === "string") {
                                var year = d.getFullYear();
                                var month = d.getMonth() + 1;
                                var day = d.getDate();
                                var hours = d.getHours();
                                var minutes = d.getMinutes();
                                var seconds = d.getSeconds();
                                var textval = column.format.replace('yyyy', year)
                                    .replace('MM', month < 10 ? "0" + month : month)
                                    .replace('dd', day < 10 ? "0" + day : day)
                                    .replace('HH', hours < 10 ? "0" + hours : hours)
                                    .replace('mm', minutes < 10 ? "0" + minutes : minutes)
                                    .replace('ss', seconds < 10 ? "0" + seconds : seconds);
                                td.innerText = textval;
                            }
                            else
                                td.innerText = v;
                            break;
                        case 'decimal':

                            if (!!column.format && !!v) {
                                if (typeof (column.format) === 'object') {
                                    if (!!column.format.fixed && typeof (column.format.fixed) === 'number') {
                                        var num = parseFloat(v);
                                        td.innerHTML = num.toFixed(2);
                                    } else {
                                        td.innerText = v;
                                    }

                                } else if (typeof (column.format) === 'string') {
                                    switch (column.format) {
                                        case 'money':
                                            td.innerHTML = parseFloat(v).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,') + "$";
                                            break;
                                        default:
                                            td.innerText = v;
                                            break;
                                    }
                                }
                            } else {
                                td.innerText = v;
                            }


                            break;
                        default:
                            td.innerText = v;
                            break;
                    }
                }
                if (column.hidden === true) {
                    td.style.display = 'none';
                }
                tr.append(td);


            }
            return tr;
        };
        const loadData = function (tablebody) {
            $(tablebody).empty();
            var sort = null;
            var scs = $('#' + table_uid + ' > thead > tr > th[sortstate]');
            var sortedColumn = scs.length > 0 ? scs[0] : null;
            if (sortedColumn !== null) {
                sort = {
                    field: $(sortedColumn).attr('sortfield'),
                    type: $(sortedColumn).attr('sortstate')
                }
            }
            datasource.fetch(function (records) {
                for (var j = 0; j < records.length; j++) {
                    var record = records[j];
                    var tr = buildTableRow(record);
                    tablebody.append(tr);

                }
                setupTableheight();
                refreshFooter();
            }, sort,paging);
        };
        const refreshFooter = function () {
            if (!!paging) {
                var footer = that.children('div.officeui-table-footer');
                footer.length > 0 && footer.replaceWith(buildFooter());
                footer.length == 0 && that.append(buildFooter());
                
            }
        };
        const refreshData = function () {
            loadData($('#' + table_uid + ' > tbody')[0]);
            //refreshFooter();
        };

        const buildBody = function() {

            var tablebody = document.createElement('tbody');
            loadData(tablebody);
            //datasource.fetch(function (records) {
            //    for (var j = 0; j < records.length; j++) {
            //        var record = records[j];
            //        var tr = buildTableRow(record);
            //        tablebody.append(tr);
            //    }

            //});
            return tablebody;
        }

        const buildTable = function () {
            var table = document.createElement('table');
            table.className = selectable ? "ms-Table ms-Table--selectable" : "ms-Table";
            table.setAttribute('uid', table_uid);
            table.id = table_uid;


            table.append(buildHeader());

            table.append(buildBody());
            return table;
        }
        const tableWrapper = function (table) {
            var wrapDiv = document.createElement('div');
            wrapDiv.className = 'officeuiExtension-table-wrapper';

            if (!!config.minheight)
                wrapDiv.style.minHeight = config.minheight;
            if (!!config.maxheight)
                wrapDiv.style.maxheight = config.maxheight;
            if (!!config.height)
                wrapDiv.style.height = config.height;
           
            wrapDiv.append(table);
            return wrapDiv;
        };
        const setupTableheight = function () {
            if (!!paging && paging.size) {
                var wd = that.children('div.officeuiExtension-table-wrapper');
                var header_h = wd.children('table').children('thead').outerHeight(true);
                rowHeight = $('#' + table_uid + ' > tbody > tr:first').outerHeight(true);
                var h = ((paging.size * rowHeight) + header_h) + 'px';
                wd.css('height', h);
                // wrapDiv.style.height = (paging.size * rowHeight) + 'px';
            }
        }
        const buildFooter = function () {
            const createFooterSpan = function (arg) {
                var s = document.createElement('span');
                s.className = !!arg && typeof (arg) === 'string' ? "officeui-table-pagingIcon" : "";
                if (!!arg && typeof (arg) === 'string') {
                    var i = document.createElement('i');
                    i.className = arg;
                    s.append(i);
                } else if (!!arg && typeof (arg) === 'number') {
                    s.innerText = arg;
                }
                return s;
                
            };
            var footerDiv = document.createElement('div');
            footerDiv.className = 'officeui-table-footer';
            datasource.getCount(function (count) {
                if (paging.displayCount == true) {
                    var divCounter = document.createElement('div');
                    divCounter.className = "ms-Grid-col ms-sm6";
                    var counter = document.createElement('span');
                    if (!!paging.size) {
                        var indexCount = (paging.currentPage || 1) > 1 ? (((paging.currentPage || 1) - 1) * paging.size) + datasource.Data.length : datasource.Data.length;
                        counter.innerText = indexCount + " out of " + count + " records";
                    }
                    else
                        counter.innerText = count + " records";
                    divCounter.append(counter);
                    footerDiv.append(divCounter);

                }
                if (!!paging.size && paging.size > 0) {

                    if (!paging.currentPage) {
                        paging.currentPage = 1;
                    }
                    var pagingDiv = document.createElement('div');
                    pagingDiv.className = "ms-Grid-col ms-sm6";
                    pagingDiv.style.textAlign = "right";
                   // pagingDiv.style.color = 'lightgray';
                    var firstPage = createFooterSpan("ms-Icon ms-Icon--ChevronLeftEnd6");

                    
                    if (paging.currentPage === 1)
                        firstPage.className = 'officeui-table-pagingIcon-disabled';
                    else {
                        firstPage.addEventListener('click', function (e) {
                            if (datasource.getDirtyRows().length > 0) {
                                alert('Save changes before change page');
                                return;
                            }
                            paging.currentPage = 1;
                            refreshData();

                        });
                    }
                    var prevPage = createFooterSpan("ms-Icon ms-Icon--ChevronLeftSmall");

                   
                    if (paging.currentPage === 1)
                        prevPage.className = 'officeui-table-pagingIcon-disabled';
                    else {
                        prevPage.addEventListener('click', function (e) {
                            if (datasource.getDirtyRows().length > 0) {
                                alert('Save changes before change page');
                                return;
                            }
                            if (paging.currentPage > 1) {
                                paging.currentPage = paging.currentPage - 1;
                                refreshData();
                            }

                        });
                    }

                    var pageNumber = createFooterSpan(paging.currentPage);
                    pageNumber.style.marginRight = "4px";
                    var nextPage = createFooterSpan("ms-Icon ms-Icon--ChevronRightSmall");
                    
                    var LastPage = createFooterSpan("ms-Icon ms-Icon--ChevronRightEnd6");
                    
                    if (count <= (paging.size * paging.currentPage)) {
                        nextPage.className = 'officeui-table-pagingIcon-disabled';
                        LastPage.className = 'officeui-table-pagingIcon-disabled';
                    } else {
                        nextPage.addEventListener('click', function (e) {
                            if (datasource.getDirtyRows().length > 0) {
                                alert('Save changes before change page');
                                return;
                            }
                            if (count > (paging.size * paging.currentPage)) {
                                paging.currentPage = paging.currentPage + 1;
                                refreshData();
                            }

                        });
                        LastPage.addEventListener('click', function (e) {
                            if (datasource.getDirtyRows().length > 0) {
                                alert('Save changes before change page');
                                return;
                            }
                            if (count > (paging.size * paging.currentPage)) {
                                paging.currentPage = Math.ceil(count / paging.size);
                                refreshData();
                            }

                        });
                    }


                    pagingDiv.append(firstPage);
                    pagingDiv.append(prevPage);
                    pagingDiv.append(pageNumber);
                    pagingDiv.append(nextPage);
                    pagingDiv.append(LastPage);
                    footerDiv.append(pagingDiv);
                }
            });
           
           
             

            return footerDiv;
        };

        var t = buildTable();
        this.append(tableWrapper(t));
       
       
       /* if (!!paging) {
            var footer = buildFooter();
            this.append(footer);
        }*/

       
        const getSelectedRows = function () {
            var rows = $('#' + table_uid + ' > tbody > tr.is-selected');
            var srows = rows.map(function (_i) {
                var uid = $(rows[_i]).attr('uid');
                var srow = datasource.Data.filter(function (d) { return d.uid === uid; })[0];
                return srow;
            });
            return srows;
        };
        const addRow = function () {
            var tbody = $("#" + table_uid).children('tbody');
            var r = rowFactory();
            tbody.prepend(buildEditableRow(r));
        };
        const deleteRow = function (arg) {
            var datarow = null;
            if (typeof (arg) === 'string') {
                datarow = datasource.Data.filter(function (d) { return d[schema.key] === arg; });
            }
            else {
                datarow = arg;
            }
            $('#' + table_uid + ' > tbody > tr[uid="' + datarow.uid + '"]').remove();
            if (datarow.mode !== 'new') {
                datarow.mode = 'removed';
                
            } else {
                var idx = datasource.Data.indexOf(datarow);
                datasource.Data.splice(idx, 1);
            }
            showDirty();
        };
       
        if (!!commandbar) {
           
            var bar_container = document.createElement('div');
            bar_container.className = "ms-CommandBar";
            

            var bar_mainarea = document.createElement('div');
            bar_mainarea.className = "ms-CommandBar-mainArea";

            for (var i = 0; i < commandbar.length; i++) {
                var btnConfig = commandbar[i];

                var btn_container = document.createElement('div');
                btn_container.className = "ms-CommandButton";
                
                var btn = document.createElement('a');
                btn.className = "ms-CommandButton-button";
                var icon = "";
                switch (btnConfig.type) {
                    case 'custom':

                        icon = btnConfig.icon;
                        btn.addEventListener('click', function (e) {
                            btnConfig.onClick(e);
                        });
                        break;
                    case "create"://Addrow

                        icon = "ms-Icon ms-Icon--Add";
                        btn.addEventListener('click', function (e) {
                            addRow();
                        });
                        break;
                    case "delete"://delete selected
                        icon = "ms-Icon ms-Icon--Delete";
                        btn.addEventListener('click', function (e) {
                            var datarows = getSelectedRows();
                            for (var i = 0; i < datarows.length; i++) {
                                deleteRow(datarows[i]);
                            }
                        });
                        break;
                    case "refresh":
                        icon = "ms-Icon ms-Icon--Refresh";
                        btn.addEventListener('click', function (e) {
                            refreshData();
                        });
                        break;
                    case "save": //save new row and existing rows
                        btnConfig.label = "Saved";
                        icon = "ms-Icon ms-Icon--Save";
                        btn.id = table_uid+'_save';
                        btn.addEventListener('click', function (e) {
                            showSaving(false);
                            datasource.saveChanges(function (r, state,isLast) {
                                var tr = null;
                                switch (state) {
                                    case 'edited':
                                    case 'new':
                                        tr = $('#' + table_uid + ' > tbody > tr[uid="' + r.uid + '"]');
                                        tr.removeAttr('mode');
                                        tr.attr('dataid', r[schema.key]);
                                        var ntr = buildTableRow(r);
                                        tr.replaceWith(ntr);
                                        break;
                                    case 'removed':
                                        break;
                                    default:
                                        break;
                                }
                                if (isLast) {
                                    showSaving(true);
                                    refreshFooter();
                                }


                            });
                        });
                        btn_container.style.float = 'right';
                        break;
                    default://TODO
                        break;
                }

                btn.innerHTML = '<span class="ms-CommandButton-icon ms-fontColor-themePrimary"><i class="' + icon + '"></i>' + (!!btnConfig.label ? '</span><span class="ms-CommandButton-label">' + btnConfig.label+'</span>': '');
                btn.href = "javascript:void(0)";
                btn_container.append(btn);
                bar_mainarea.append(btn_container);
            }
            bar_container.append(bar_mainarea);

        }
        this.prepend(bar_container);
        

        var that = this;

       
        var grid = {
            element: that,
            officeEl: new fabric['Table'](document.getElementById(table_uid)),
            dataSource: datasource,
            getSelectedRows: getSelectedRows,
            refresh: refreshData,
            addRow: addRow,
            deleteROw: deleteRow,
            updateRow: function (uid) {
    
            }
        };
        return grid;
        

    };

    $.fn.OfficeUIDropdown = function (config) {
        /* 
         * <div class="ms-Dropdown" tabindex="0">
              <label class="ms-Label">Dropdown label</label>
              <i class="ms-Dropdown-caretDown ms-Icon ms-Icon--ChevronDown"></i>
              <select class="ms-Dropdown-select">
                <option>Choose a sound&amp;hellip;</option>
                <option>Dog barking</option>
                <option>Wind blowing</option>
                <option>Duck quacking</option>
                <option>Cow mooing</option>
              </select>
            </div>
         * */
        function uuid() {
            return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
                var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
                return v.toString(16);
            });
        }
        var datasource = config.datasource;
        var labelkey = config.labelkey;
        var valuekey = config.valuekey;
        var selectedvalue = config.selectedvalue;
        var c_uid = uuid();

        var container = document.createElement('div');
        container.className = 'ms-Dropdown';
        container.tabIndex = 0;
        container.id = c_uid;

        var chevron = document.createElement("i");
        chevron.className = "ms-Dropdown-caretDown ms-Icon ms-Icon--ChevronDown";
        chevron.style.bottom = "5px";
        container.append(chevron);

        var select = document.createElement('select');
        select.className = "ms-Dropdown-select";
        select.addEventListener('change', function (e) {
            if (!!config.onChange) {
                config.onChange(e);
            }
        });
        datasource.fetch(function (records) {
            for (var i = 0; i < records.length; i++) {
                var r = records[i];

                var opt = document.createElement('option');
                opt.value = r[valuekey] || "";

                if (!!selectedvalue && opt.value === selectedvalue) {
                    opt.selected = true;
                }

                opt.innerText = r[labelkey] || "";
                select.append(opt);

            }
        });
        var that = this;
        container.append(select);
        this.append(container);
       
        var dropdown = {
            element: that,
            officeEl: new fabric['Dropdown'](container),
            select: function (value) {
                $(this.element).children("div > select > option[value='" + value + "']").attr('selected', 'selected');
            },
            getValue: function () {
                var x = $(this.element).children('div');
                var y = x.children('select');
                var opt = y.children('option[value="'+y.val()+'"]');
                
                return { label:opt.text(),value:opt.val()};
            }

        };
        return dropdown;


    };
    

}(jQuery));

