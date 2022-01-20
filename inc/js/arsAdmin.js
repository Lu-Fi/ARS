// ADMIN MENU FILES AND FUNCTIONS
const _ars_adm_menu = {
    adm_assignments: {
        file: './inc/html/adm_assignments.html',
        done: function () {
            arsAdmAssignmentInit();
        }
    },
    adm_accounts: {
        file: './inc/html/adm_accounts.html',
        done: function () {
            arsAdmAccountsInit();
        }
    },
    adm_audit: {
        file: './inc/html/adm_audit.html',
        done: function () {
            arsAdmAuditInit();
        }
    }
}

// LOAD ADMIN MENU ENTRY
function loadAdmMenuEntry(e, f = null) {

    if ($('#arsBusy').is(':visible')) { return; }

    if (typeof e !== 'undefined' && typeof _ars_adm_menu[e] !== 'undefined') {

        $('#arsCw').load(_ars_adm_menu[e].file, function () {

            _ars_adm_menu[e].done();
            if (typeof f == 'function') {
                f();
            }
        });
    }
}

// MAIN PAGE INIT
(function ($, window, document) {

    $(function () {

        $('#arsNavAdminMenu .nav-link').click(function () {

            $('#arsNav .nav-link').removeClass('active');
            $(this).addClass('active');

            loadAdmMenuEntry(
                $(this).attr('data-menu'));
        });
    });
}(window.jQuery, window, document));

function arsAdmAccountsInit() {

    arsCwData = {};

    arsCwData.Button = $('#tplButton').html();
    arsCwData.Switch = $('#tplSwitch').html();
    arsCwData.tplRemove = $('#tplConfirmRemove').html();
    arsCwData.tplSearch = $('#tplSearch').html();

    arsResize = function () {

        /*
            .description
            Resize function for assignments menu entry
        */

        $('.arsCsw .card').css("height",
            $('.arsCsw.accounts').first().height() - 100);

        $('.arsCsw .dataTables_scrollBody').css("max-height",
            $('.arsCsw .card').first().height() - 145);

        if (typeof arsCwData.UsrTbl != 'undefined') {

            arsCwData.UsrTbl.columns.adjust();
        }
    }

    arsCwData.initSearch = function (i, m, c = 3, f = null) {

        /*
            .description
            Initialize typeahead autocomplete to search
            for user and or group

            .parameter
            i = Input object to bind typeahed
            m = maxHeigh for autocomplete list
            c = object class in search, 0 = user or group, 1 = user, 2 = group
            f = callbeack function on select
        */

        $(i).typeahead({
            minLength: 3,
            items: 20,
            source: function (q, r) {

                arsResize();

                if (arsCwData.thTimeout) {

                    clearTimeout(arsCwData.thTimeout);
                }

                arsCwData.thTimeout = setTimeout(function () {

                    $.ajax({
                        data: {
                            'ebbe0df1-2bcd-4251-895e-aea33092716e': c, 's': q
                        },
                        success: function (data) {

                            r(data);
                        }
                    });
                }, 400);
            },
            matcher: function (item) {

                return item;
            },
            sorter: function (items) {

                return items;
            },
            displayText: function (item) {

                return item;
            },
            updater: function (item) {

                $(i).attr(
                    "data-sid", item.sid).attr(
                        "data-type", item.type);

                if (item.displayname) {

                    return item.displayname
                }
                return item.name;
            },
            highlighter: function (item) {

                if (item.type == '2') {

                    return '<div class="arsThEntry border-bottom"><i class="bi-people"></i><span>' + item.name + '</span><br /><span>Description:<span>' + item.description + '</span></div>';
                } else {

                    var n = item.name;

                    if (item.displayname) {

                        n = item.displayname;

                        if (item.displayname.indexOf(item.name) < 0) {

                            n += ' (' + item.name + ')';
                        }
                    }
                    return '<div class="arsThEntry"><i class="bi-person"></i><span>' + n + '</span><br /><span>Company:<span>' + item.company + '</span>,Department:<span>' + item.department + '</span></div>';
                }
            },
            afterSelect: function (item) {

                if (typeof f == 'function') {

                    f();
                }
            },
            maxHeight: m
        });
    }

    arsCwData.LoadTableData = function () {

        /*
            .description
            Load Account data and fill table
        */

        arsBusy();

        $.ajax({

            data: { 'a6dc5917-0837-4eb8-bede-2c150969b627': '' }
        }).done(function (data) {

            if (data.length > 0) {

                arsCwData.tableData = data.Users;

                arsCwData.UsrTbl.clear();
                arsCwData.UsrTbl.rows.add(data);
                arsCwData.UsrTbl.draw();

                arsBusy(true);
            }
        }).fail(function () {

            arsBusy(true);
        });
    }

    arsCwData.UsrTbl = $('#UsrTbl').DataTable({

        scrollY: '180px',
        scrollCollapse: true,
        paging: false,
        columnDefs: [
            {
                targets: 0,
                data: "DisplayName",
                createdCell: function (td, cellData, rowData, row, col) {

                    $(td).text(UsernameFromRow(rowData));
                }
            },
            {
                targets: 1,
                data: "Description"
            },
            {
                targets: 2,
                data: null,
                defaultContent: "",
                createdCell: function (td, cellData, rowData, row, col) {

                    $(td).append(
                        $('<div class="form-check form-switch"></div>').append(
                            $('<input class="form-check-input" type="checkbox">').on('click', { rowData: rowData },
                                function (event) {
                                    event.stopPropagation();
                                    arsCwData.toggleFlag(rowData.SID, 2);
                                }
                            ).prop('checked', ((rowData.flags & 1) == 1) ? true : false)
                        )                 
                    );
                },
                width: "12%"
            },
            {
                targets: 3,
                data: null,
                defaultContent: "",
                createdCell: function (td, cellData, rowData, row, col) {

                    $(td).append(
                        $('<div class="form-check form-switch"></div>').append(
                            $('<input class="form-check-input" type="checkbox">').on('click', { rowData: rowData },
                                function (event) {
                                    event.stopPropagation();
                                    arsCwData.toggleFlag(rowData.SID, 1);
                                }
                            ).prop('checked', ((rowData.flags & 256) == 256) ? true : false)
                        )
                    );
                },
                width: "12%"
            },
            {
                targets: 4,
                data: null,
                defaultContent: "",
                createdCell: function (td, cellData, rowData, row, col) {

                    $(td).append(
                        $('<div class="input-group"></div>').append(
                            $(arsCwData.Button).click(
                                function (event) {
                                    event.stopPropagation();
                                    arsCwData.removeUserPopup(rowData.SID);
                                }
                            ).text('Remove')
                        )
                    );
                },
                width: "1%"
            }
        ]
    }).on('click', 'tr', function (event) {

        var rowData = $(event.delegateTarget).DataTable().row(this).data();

        $('#arsNav .nav-link').removeClass('active');
        $('#arsNav .nav-link[data-menu="adm_assignments"]').addClass('active');

        loadAdmMenuEntry('adm_assignments', function () { arsCwData.loadSid(rowData.SID, 1) });
    });;

    arsCwData.removeUserPopup = function (uSID) {

        arsPopup("Remove account ",
            arsCwData.tplRemove,
            [
                { action: 'cancel', callback: function () { $('#arsModal').modal('hide'); } },
                { action: 'yes', callback: function () { arsCwData.removeUser(uSID); } }
            ]);
    }

    arsCwData.removeUser = function (uSID) {

        $.ajax({
            data: {
                'c5fe4461-3c40-4c5c-a61a-8afbde66485b': uSID
            },
            success: function (data) {

                arsCwData.LoadTableData();
                $('#arsModal').modal('hide');
            },
            fail: function () {

                $('#arsModal').modal('hide');
            }
        });
    }

    arsCwData.addUser = function () {

        /*
            .description
            Send add account request

            .send parameter
            SID to add
        */

        var e = $('#arsSearchAdd');
        var s = e.attr('data-sid');

        if (s) {

            $.ajax({
                data: {
                    '2db55222-a29e-4656-ac82-e63fe34a2ea9': s
                },
                success: function (data) {

                    arsCwData.LoadTableData();
                    $('#arsModal').modal('hide');
                }
            });
        } else {

            $('.modal-body').addClass("alert-danger", 1000).
                removeClass("alert-danger", 1000)

            var f = $('.modal-footer');
            f.find('.alert').remove();
            f.prepend('<div class="alert alert-danger" role="alert">Please select entry from dropdown.</div>');
        }
    }

    arsCwData.toggleFlag = function (uSID, f) {

        $.ajax({
            data: {
                '0bd9da76-7ed3-4eb9-a49c-0b4ba4b1bc14': uSID, 'f':f
            },
            success: function (data) {

                arsCwData.LoadTableData();
            }
        });
    }

    arsCwData.selectObject = function () {

        /*
            .description
            Create a Dialog with Directory Search Functionality 
            to add a new User
        */

        arsPopup("Select account",
            arsCwData.tplSearch,
            [
                { action: 'cancel', callback: function () { $('#arsModal').modal('hide'); } },
                { action: 'ok', callback: function () { arsCwData.addUser(); } }
            ],
            function () {
                arsCwData.initSearch(
                    '#arsSearchAdd',
                    function () {
                        return ($('#arsCw').height() / 2);
                    }, 1);
            });
    }

    $('#btnAddUsr').click(function () {

        arsCwData.selectObject()
    })

    $('.arsAccountsInfo button').click(function () {

        arsCwData.LoadTableData();
    });

    arsResize();
    setTimeout(arsCwData.LoadTableData, 50);
}

function arsAdmAssignmentInit() {

    /*
        .description
        Initialize assignments menu entry
    */

    arsCwData = {};
    arsCwData.thTimeout = null;
    arsCwData.ResultData = null;

    arsCwData.Csw = $('.arsCsw.admin');
    arsCwData.tplUser = $('#tplUser').html();
    arsCwData.tplGroup = $('#tplGroup').html();
    arsCwData.tplSearch = $('#tplSearch').html();
    arsCwData.Button = $('#tplButton').html();
    arsCwData.tplSwitch = $('#tplSwitch').html();
    arsCwData.Advanced = $('#tplAdvancedOptions').html();
    arsCwData.PwBox = $('#tplArsPw').html();
    arsCwData.tplGroupType = $('#tplGroupType').html();

    arsResize = function () {

        /*
            .description
            Resize function for assignments menu entry
        */

        $('.typeahead').first().css('max-height',
            $('#arsCw').height() / 2);

        $('.arsCsw .card').css("height",
            ($('.arsCsw.admin').first().height() / 2) - 50);

        $('.arsCsw .dataTables_scrollBody').css("max-height",
            $('.arsCsw .card').first().height() - 130);

        if (typeof arsCwData.UsrTbl != 'undefined') {

            arsCwData.UsrTbl.columns.adjust();
        }

        if (typeof arsCwData.AssUsrTbl != 'undefined') {

            arsCwData.AssUsrTbl.columns.adjust();
        }

        if (typeof arsCwData.AssGrpTbl != 'undefined') {

            arsCwData.AssGrpTbl.columns.adjust();
        }
    }

    arsCwData.initSearch = function (i, m, c = 3, f = null) {

        /*
            .description
            Initialize typeahead autocomplete to search
            for user and or group

            .parameter
            i = Input object to bind typeahed
            m = maxHeigh for autocomplete list
            c = object class in search, 0 = user or group, 1 = user, 2 = group
            f = callbeack function on select
        */

        $(i).typeahead({
            minLength: 3,
            items: 20,
            source: function (q, r) {

                arsResize();

                if (arsCwData.thTimeout) {

                    clearTimeout(arsCwData.thTimeout);
                }

                arsCwData.thTimeout = setTimeout(function () {

                    $.ajax({
                        data: {
                            'ebbe0df1-2bcd-4251-895e-aea33092716e': c, 's': q
                        },
                        success: function (data) {

                            r(data);
                        }
                    });
                }, 400);
            },
            matcher: function (item) {

                return item;
            },
            sorter: function (items) {

                return items;
            },
            displayText: function (item) {

                return item;
            },
            updater: function (item) {

                $(i).attr(
                    "data-sid", item.sid).attr(
                        "data-type", item.type);

                if (item.displayname) {

                    return item.displayname
                }
                return item.name;
            },
            highlighter: function (item) {

                if (item.type == '2') {

                    return '<div class="arsThEntry border-bottom"><i class="bi-people"></i><span>' + item.name + '</span><br /><span>Description:<span>' + item.description + '</span></div>';
                } else {

                    var n = item.name;

                    if (item.displayname) {

                        n = item.displayname;

                        if (item.displayname.indexOf(item.name) < 0) {

                            n += ' (' + item.name + ')';
                        }
                    }
                    return '<div class="arsThEntry"><i class="bi-person"></i><span>' + n + '</span><br /><span>Company:<span>' + item.company + '</span>,Department:<span>' + item.department + '</span></div>';
                }
            },
            afterSelect: function (item) {

                if (typeof f == 'function') {

                    f();
                }
            },
            maxHeight: m
        });
    }

    arsCwData.loadSid = function ( s = '', t = 0 ) {

        /*
            .description
            Load user or group information from server 
            and initialize view

            .parameter
            s = sid to view
            t = type of sid (1 = user or 2 = group)
        */

        arsBusy();

        var o = $('#arsSearch');

        if (!s && !t) {

            s = o.attr("data-sid");
            t = o.attr("data-type");
        } else {

            o.attr("data-sid", s);
            o.attr("data-type", t);
        }

        if (s && t) {

            $.ajax({
                data: {
                    '1f87eaa6-d780-4281-a429-856672746f19': s, 't': t
                },
                success: function (data) {

                    arsCwData.ResultData = data;

                    if ('Type' in data) {

                        arsCwData.InitGroup(data);
                    } else {

                        arsCwData.InitUser(data);
                    }
                    
                    arsBusy(true);
                }
            });
        }
    }

    arsCwData.InitUser = function (data) {

        /*
            .description
            Initialize user to display
        */

        arsCwData.Csw.html('');

        var tplUser = $(arsCwData.tplUser);
        var n = data.DisplayName;

        if (!n) {

            n = data.Name;
        }

        tplUser.find('#chUserName').text(n);
        tplUser.find('#chUserInfo').html('Department: <span>' + data.Department + '</span>, Company: <span>' + data.Company + '</span>');

        tplUser.find('#btnAddAssUsr').click(function () {

            arsCwData.selectObject('Assigned User', 2);
        })

        /*
        tplUser.find('#btnAddUsr').click(function () {

            arsCwData.selectObject('User', 1);
        })
        */
        tplUser.find('#btnAddUsr').css('display', 'none');

        tplUser.find('#btnAddAssGrp').click(function () {

            arsCwData.selectObject('Group', 3);
        })

        tplUser.find('#btnAdvanced')
            .css('display', (data.AssignedUsers.length>0)?'block':'none')
            .click(function () {

                arsPopup("AdvancedOptions", arsCwData.Advanced)
            })

        tplUser.find('#btnRefresh').click(function () {

            arsCwData.loadSid();
        });

        arsCwData.Csw.append(tplUser);

        arsCwData.initUsrTbl('UsrTbl');
        arsCwData.UsrTbl.clear();
        arsCwData.UsrTbl.rows.add(data.Users);
        arsCwData.UsrTbl.draw();

        arsCwData.initUsrTbl('AssUsrTbl', 1);
        arsCwData.AssUsrTbl.clear();
        arsCwData.AssUsrTbl.rows.add(data.AssignedUsers);
        arsCwData.AssUsrTbl.draw();

        arsCwData.initGrpTbl('AssGrpTbl');
        arsCwData.AssGrpTbl.clear();
        arsCwData.AssGrpTbl.rows.add(data.AssignedGroups);
        arsCwData.AssGrpTbl.draw();

        arsResize();
    }

    arsCwData.InitGroup = function (data) {

        /*
            .description
            Initialize Group to display
        */

        arsCwData.Csw.html('');

        var tplGroup = $(arsCwData.tplGroup).clone();

        tplGroup.find('#chGroupName').text(data.Name);
        tplGroup.find('#chGroupInfo').html('Description: <span>' + data.description + '</span>');

        tplGroup.find('#btnAddAssGrp').click(function () {

            arsCwData.selectObject('Assigned User', 2);
        })

        arsCwData.Csw.append(tplGroup);

        arsCwData.initUsrTbl('UsrTbl');
        arsCwData.UsrTbl.clear();
        arsCwData.UsrTbl.rows.add(data.Users);
        arsCwData.UsrTbl.draw();

        $('.arsUserInfo button').click(function () {
            arsCwData.loadSid();
        });

        arsResize();
    }

    arsCwData.selectObject = function (t, n) {

        /*
            .description
            Create a Dialog with Directory Search Functionality 
            to add a new Assignment to User, AssignedUser or Group

            .parameter
            t = Type string used in Dialog Title
            n = Assignement type number
        */

        var c = (n > 2 ? 2 : 1); //Object class in search

        arsPopup("Select " + t,
            arsCwData.tplSearch,
            [
                { action: 'cancel', callback: function () { $('#arsModal').modal('hide'); } },
                { action: 'ok', callback: function () { arsCwData.sendAddRequest(n); } }
            ],
            function () {
                arsCwData.initSearch(
                    '#arsSearchAdd',
                    function () {
                        return ($('#arsCw').height() / 2);
                    }, c);
            });
    }

    arsCwData.sendAddRequest = function (t) {

        /*
            .description
            Send add assignement request

            .parameter
            t = Assignment type

            .send parameter
            s = SID to add
            t = Assignement type
        */

        var e = $('#arsSearchAdd');
        var s = e.attr('data-sid');

        if (s && t) {

            $.ajax({
                data: {
                    '13178d9e-1e6a-4fd3-bcd5-e13fdd8e986b': s,
                    't': t
                },
                success: function (data) {

                    if ('state' in data) {

                        if (data.state < 0) {

                            arsPopup("Access Resquest",
                                $('#tplAddResult').html(),
                                [
                                    { action: 'cancel', callback: function () { $('#arsModal').modal('hide'); } },
                                    { action: 'yes', callback: function () { arsCwData.confirmAddRequest(); } }
                                ]);
                        } else {

                            $('#arsModal').modal('hide');
                            arsCwData.loadSid();
                        }
                    }
                }
            });
        } else {

            $('.modal-body').addClass("alert-danger", 1000).
                removeClass("alert-danger", 1000)

            var f = $('.modal-footer');
            f.find('.alert').remove();
            f.prepend('<div class="alert alert-danger" role="alert">Please select entry from dropdown.</div>');
        }
    }

    arsCwData.confirmAddRequest = function () {
        $.ajax({
            data: {
                '13178d9e-1e6a-4fd3-bcd5-e13fdd8e986b': ''
            },
            success: function (data) {

                $('#arsModal').modal('hide');

                if ('state' in data) {

                    if (data.state = 0) {

                        
                    }
                }

                arsCwData.loadSid();
            }
        });
    }

    arsCwData.initUsrTbl = function (i, a = 0) {

        arsCwData[i] = $('#' + i).DataTable({

            scrollY: '180px',
            scrollCollapse: true,
            paging: false,
            columnDefs: [
                {
                    targets: 0,
                    data: "DisplayName",
                    createdCell: function (td, cellData, rowData, row, col) {

                        $(td).text(UsernameFromRow(rowData));
                    }
                },
                {
                    targets: 1,
                    data: "Description"
                },
                {
                    targets: 2,
                    data: null,
                    defaultContent: "",
                    createdCell: function (td, cellData, rowData, row, col) {

                        //if (!rowData.rSID) {

                            var g = $('<div class="input-group"></div>');

                            if (a == 1) {
                                g.append(
                                    $(arsCwData.Button).click(
                                        function (event) {
                                            event.stopPropagation();
                                            arsPopup("AdvancedOptions", arsCwData.AdvancedOptions(rowData), [], false, arsCwData.loadSid);
                                        }
                                    ).attr('title', "Advanced options").html('<i class="bi bi-gear"></i>')
                                );
                            }

                            g.append(
                                $(arsCwData.Button).click(
                                    function (event) {
                                        event.stopPropagation();
                                        arsCwData.removeAssignment(rowData.aID);
                                    }
                                ).prop('disabled', (rowData.source != 1) ? true : false).text('Remove')
                            );

                            $(td).append(g);
                        //}
                    },
                    width: "1%"
                }
            ]
        }).on('click', 'tr', function (event) {

            var rowData = $(event.delegateTarget).DataTable().row(this).data();
            arsCwData.loadSid(rowData.SID, 1);
        });
    }

    arsCwData.showPwd = function () {

        if (arsCwData.Timeout) {

            clearTimeout(arsCwData.Timeout);
        }

        $('#arsPwIg input').attr('type', 'text');
        $('#arsPwIg i').removeClass("bi-eye-slash").addClass("bi-eye");

        arsCwData.Timeout = setTimeout(function () {
            arsCwData.hidePwd()
        }, 10000)
    }

    arsCwData.hidePwd = function () {

        if (arsCwData.Timeout) {

            clearTimeout(arsCwData.Timeout);
        }

        $('#arsPwIg input').attr('type', 'password');
        $('#arsPwIg i').removeClass("bi-eye").addClass("bi-eye-slash");
    }

    arsCwData.AdvancedOptions = function (rowData) {

        var r = $(arsCwData.Advanced);

        r.find("[data-option]").each(function () {

            var n = $(this);
            var o = n.attr('data-option');

            n.click(function () {

                arsCwData.toggleFlag(rowData.aID, o, n);
            })

            if ((rowData.flags & o) == o) {

                n.prop('checked', true);
            }
        });

        r.find("button").on('click', function () {

            var n = $(this);

            $.ajax({

                data: { '772005f4-edb5-4fb8-a723-2dc00af2cc09': rowData.aID }
            }).done(function (data) {

                if (data.action == 'popup') {

                    var p = $(arsCwData.PwBox);

                    p.find("button").on('click', function (event) {

                        event.preventDefault();

                        var n = $('#arsPwIg input');

                        if (n.attr("type") == "text") {

                            arsCwData.hidePwd()
                        } else if (n.attr("type") == "password") {

                            arsCwData.showPwd()
                        }
                    });

                    p.find('input').val(data[rowData.aID]);

                    arsPopup('Reset password', p, [{ action: 'close', callback: function () { $('#arsModal').modal('hide'); } }]);
                }
            })
        })

        return r;
    }

    arsCwData.initGrpTbl = function (i) {

        arsCwData[i] = $('#' + i).DataTable({

            scrollY: '180px',
            scrollCollapse: true,
            paging: false,
            columnDefs: [
                {
                    targets: 0,
                    data: "Name"
                },
                {
                    targets: 1,
                    data: "Description"
                },
                {
                    targets: 2,
                    data: null,
                    defaultContent: "",
                    createdCell: function (td, cellData, rowData, row, col) {

                        $(td).append(
                            $('<div class="input-group"></div>').append(
                                $(arsCwData.Button).click(
                                    function (event) {
                                        event.stopPropagation();
                                        arsCwData.removeAssignment(rowData.aID);
                                    }
                                ).prop('disabled', (rowData.source != 1) ? true : false).text('Remove')
                            )
                        );
                    },
                    width: "1%"
                }
            ]
        }).on('click', 'tr', function (event) {

            var rowData = $(event.delegateTarget).DataTable().row(this).data();
            arsCwData.loadSid(rowData.SID, 2);
        });
    }

    arsCwData.removeAssignment = function (aID) {

        $.ajax({
            data: {
                'ff90792a-4c63-445d-b981-ee80e6263891': aID
            },
            success: function (data) {

                arsCwData.loadSid();
            }
        });
    }

    arsCwData.toggleFlag = function (aID, f, n) {

        $.ajax({

            data: { 'f3e7708f-e9b4-4219-8557-d5a730bbf70c': aID, 'f': f }
        }).done(function (data) {

            if (!('state' in data) || data.state != 1) {

                n.prop('checked', n.is(':checked') ? false : true)
            }
        })
    }

    arsCwData.initSearch(
        '#arsSearch',
        function () {
            return ($('#arsCw').height() / 2);
        }, 3, arsCwData.loadSid);

    arsResize();
}

function arsAdmAuditInit() {

    arsCwData = {};

    arsResize = function () {

        arsCwData.AuditTable.columns.adjust();
    }

    arsCwData.AuditTable = $('#arsAdmAuditLog').DataTable({
        processing: true,
        serverSide: true,
        ajax: {
            url: "./inc/asp/data.aspx",
            type: "POST",
            data: { 'cdbb6e5e-48df-419c-8f7d-9a4bc8df25b7': '' }
        },
        scrollResize: true,
        scrollX: true,
        scrollY: 100,
        scrollCollapse: true,
        paging: true,
        lengthChange: false,
        pageLength: 50,
        columnDefs: [
            {
                targets: 0,
                data: "Time",
                render: $.fn.dataTable.render.moment('YYYY-MM-DD HH:mm:ss.SSS', 'DD/MM/YYYY HH:mm:ss.SSS')
            },
            {
                targets: 1,
                data: "Account"
            },
            {
                targets: 2,
                data: "Source"
            },
            {
                targets: 3,
                data: "AssignedAccount",
            },
            {
                targets: 4,
                data: "Action"
            },
            {
                targets: 5,
                data: "Result",
            }
        ]
    });

    $('.dataTables_filter input').attr('maxlength', 64);

    arsResize();
}