var _ars_busyTimeout = null;// TIMEOUT FOR BUSY INDICATOR
var _ars_rt = null;         // ARS GLOBAL RESIZE TIMEOUT
var arsResize = null;       // RESIZE FUNCTION 
var arsCwData = {};         // DATA HOLDER 
const _ars_menu = {

    /*
        .description
        menu entry target and callback structure
    */

    groups: {
        file: './inc/html/groups.html',
        done: function () {
            arsGroupsInit();
        }
    },
    advanced: {
        file: './inc/html/advanced.html',
        done: function () {
            arsAdvancedInit();
        }
    },audit: {
        file: './inc/html/audit.html',
        done: function () {
            arsAuditInit();
        }
    }
}

$(window).resize(function () {

    clearTimeout(_ars_rt);

    if (typeof arsResize === 'function') {

        _ars_rt = setTimeout(arsResize, 50);
    }
});

function loadMenuEntry(e) {

    /*
        .description
        load menu entry by using _ars_menu structure
    */

    if ($('#arsBusy').is(':visible')) { return; }

    if (typeof e !== 'undefined' && typeof _ars_menu[e] !== 'undefined') {

        $('#arsCw').load(_ars_menu[e].file, function () {

            _ars_menu[e].done();
        });
    }
}

function arsBusy(done = false) {

    /*
        .description
        display full page overlay spinner

        .parameter
        done = set  overlay spinner hidden
    */

    clearTimeout(_ars_busyTimeout);

    if (done == false) {

        _ars_busyTimeout = setTimeout(function () {

            $('#arsBusy').fadeIn(200);
        }, 500);
    } else {

        _ars_busyTimeout = setTimeout(function () {

            $('#arsBusy').fadeOut(200);
        }, 10);
    }
}

function getToJson(lStrGet) {

    var o = {};

    if (lStrGet !== undefined) {

        var d = lStrGet.split("&");

        for (var k in d) {

            o[d[k].split("=")[0]] = d[k].split("=")[1];
        }
    }
    return o;
}

function arsPopup(title, body, buttons, init, deinit) {

    /*
        .description
        Create a modal popup window

        .parameter
        title = Window title
        body = body content as html
        buttons = array of buttons to add
                  [{action:<button text>,callback:<callback function>}]
        init = function which will be called after dialog render
    */

    var arsModal = $('#arsModal');

    f = function () {

        arsModal.find(".modal-title").first().text(title);
        arsModal.find(".modal-body").first().html(body);

        var footer = arsModal.find(".modal-footer").first();
        footer.empty();

        if (typeof buttons == 'object' ) {

            for (var i = 0; i < buttons.length; i++) {

                var btn = $('<button type="button" class="btn btn-primary">' + buttons[i].action + '</button>').click(

                    buttons[i].callback
                )

                footer.append(btn);
            }
        }

        if (typeof init == 'function' ) {

            init()
        }

        if (typeof deinit == 'function') {

            if (! arsModal.is(":visible")) {

                arsModal.on('hidden.bs.modal', function (e) {

                    $(this).off('hidden.bs.modal');
                    deinit();
                })
            }
        }

        arsModal.modal('show');
    }

    if (arsModal.is(":visible")) {

        arsModal.on('hidden.bs.modal', { f: f }, function (e) {

            $(this).off('hidden.bs.modal');
            e.data.f();
        })

        arsModal.modal('hide');
    } else {

        f();
    }
}

function UsernameFromRow(rowData) {
    
    var d = '';

    if (rowData.Name) {

        if (!rowData.DisplayName) {

            return rowData.Name;
        } 

        if (rowData.DisplayName.indexOf('(' + rowData.Name + ')') === -1) {

            d = rowData.Name;
        }

        if (rowData.DisplayName && d) {

            d = rowData.DisplayName + ' (' + d + ')';
        } else {

            d = rowData.DisplayName;
        }
    }
    return d;
}

(function ($, window, document) {

    /*
    .description
    document ready function
    initialize some base functionality
    */

    $(function () {

        jQuery.ajaxSetup({
            cache: false,
            async: true,
            method: "POST",
            url: "./inc/asp/data.aspx",
            beforeSend: function (xhr, settings) {

                var lReqDat = getToJson(settings.data);

                if ('p' in lReqDat) {

                    window.history.pushState(settings.data, null, './#' + lReqDat.p);
                }

                xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
                xhr.setRequestHeader('X-Referrer', 'ccArs');
                return xhr;
            },
            complete: function (data) {

                if(typeof data != 'undefined' && 'responseJSON' in data && 'ars' in data.responseJSON) {

                    if (typeof data.responseJSON.ars == 'number') {

                        if (data.responseJSON.ars == -1) {

                            window.location.reload(false);
                        }
                    } else if (typeof data.responseJSON.ars == 'string') {

                        console.log(data.responseJSON.ars);
                    }
                }
            }
        });

        $('#btnAction').click(function () {

            $('#authForm').submit()
        });

        var navLinks = $('#arsNavPersonalMenu .nav-link');
        if (navLinks.length > 0) {

            navLinks.click(function () {

                $('#arsNav .nav-link').removeClass('active');
                $(this).addClass('active');

                loadMenuEntry(
                    $(this).attr('data-menu'));
            });

            loadMenuEntry(
                $('#arsNavPersonalMenu a.active:first').attr("data-menu"));           
        }

        var b = $("button[data-azexpts]");
        if (typeof b.length != 'undefined') {

            var d = new Date(b.attr("data-azexpts") * 1000);

            b.parent().attr("Title", "Azure authenticated\r\nCookie expiration: " +d.toString());
            b.children("span").last().text(d.toLocaleTimeString());
        }
    });
}(window.jQuery, window, document));

function arsGroupsInit() {

    arsCwData = {};
    arsCwData.tableData = [];
    arsCwData.activeUser = null;

    arsCwData.Button = $('#arsButton').html();
    arsCwData.tplSwitch = $('#tplSwitch').html();
    arsCwData.TimeSelectHtml = $('#arsTimeSelect').html();
    arsCwData.UserName = $('#chUserName');
    arsCwData.UserInfo = $('#chUserInfo');

    arsResize = function() {

        $('.arsCsw .dataTables_scrollBody').css("max-height",
            $('.arsCsw .arsManagedAccounts').first().height() - 160);

        arsCwData.AssUsrTable.columns.adjust();
        arsCwData.AssGrpTable.columns.adjust();
    }

    arsCwData.LoadTableData = function (guid = '9f35e966-3c8d-43ec-923f-1346246183f9') {

        arsBusy();

        $.ajax({

            data: { [guid]: '' }
        }).done(function (data) {

            if ('Name' in data) {

                arsCwData.UserName.text(data.DisplayName);
                arsCwData.UserInfo.html('Department: <span>' + data.Department + '</span>, Company: <span>' + data.Company + '</span>');

                arsCwData.tableData = data.AssignedUsers;
                arsCwData.FillUserTable();
                arsCwData.SelectUser();

                arsBusy(true);

                if (typeof arsCwData.lastResult != 'undefined' ) {

                    for (aID in arsCwData.lastResult) {

                        if ('AddMemberToGroup' in arsCwData.lastResult[aID] ||
                            'RemoveMemberFromGroup' in arsCwData.lastResult[aID]) {

                            var r = 'danger';

                            arsCwData.AssGrpTable.rows(function (idx, data, node) {

                                if (data.aID === aID) {
                                    
                                    if (arsCwData.lastResult[aID]?.AddMemberToGroup?.result == 1 ||
                                        arsCwData.lastResult[aID]?.RemoveMemberFromGroup?.result == 1) {

                                        r = 'success';
                                    }

                                    $(node).
                                        addClass("alert-" + r, 1500).
                                            removeClass("alert-" + r, 1500)
                                }
                            })
                        }
                    }

                    /*
                    for (var i = 0; i < arsCwData.AssGrpTable.rows().count(); i++) {

                        var r = 'danger';

                        for (aID in arsCwData.lastResult) {

                            if (aID == arsCwData.AssGrpTable.row(i).data().aID) {

                                if (arsCwData.lastResult[aID]?.AddMemberToGroup?.result == 1 || arsCwData.lastResult[aID]?.RemoveMemberFromGroup?.result == 1 ) {

                                    r = 'success';
                                }

                                $(arsCwData.AssGrpTable.row(i).node()).
                                    addClass("alert-" + r, 1500).
                                        removeClass("alert-" + r, 1500)
                            }
                        }
                    }
                    */
                }
            } else {

                arsBusy(true);
            }

            arsCwData.lastResult = null;
        });
    }

    arsCwData.FillUserTable = function() {

        var dt = arsCwData.AssUsrTable;

        dt.clear();
        dt.rows.add(arsCwData.tableData);
        dt.draw();
    }

    arsCwData.FillGroupTable = function(rowData) {

        $('#arsActiveUserName').text("'" + rowData.DisplayName + "'");

        var dt = arsCwData.AssGrpTable;

        dt.clear();
        dt.rows.add(rowData.AssignedGroups);
        dt.draw();
    }

    arsCwData.SelectUser = function( aID = null ) {

        var entryFound = false;

        if (aID == null) {

            aID = arsCwData.activeUser;
        } else {

            arsCwData.activeUser = aID;
        }

        if (typeof aID != 'undefined' && aID != null) {

            arsCwData.AssUsrTable.rows().every(

                function (rowIdx, tableLoop, rowLoop) {
                    
                    var rowData = this.data();

                    if (rowData.aID == aID) {

                        arsCwData.FillGroupTable(rowData);
                        entryFound = 1;
                    }
                }
            );
        }

        if (!entryFound) {

            arsCwData.activeUser = null;
            arsCwData.AssGrpTable.clear().draw();
            $('#arsActiveUserName').text("<select Managed Account>");
        }
    }

    arsCwData.RequestGroup = function (aID, s) {

        arsBusy();

        $.ajax({

            data: { '528e8218-b284-4475-a59e-9bad1f180208': aID, 't': s }
        }).done(function (data) {

            arsCwData.lastResult = data;
            arsCwData.LoadTableData();
            arsBusy(true);
        });
    }

    arsCwData.RemoveGroup = function (aID) {

        arsBusy();

        $.ajax({

            data: { 'cc5ddb45-c324-4208-a22a-be6f809edcd9': aID }
        }).done(function (data) {

            arsCwData.lastResult = data;
            arsCwData.LoadTableData();
            arsBusy(true);
        });
    }

    arsCwData.RequestBaseGroups = function(aID, s)  {

        arsBusy();

        baseGroups = [];
        arsCwData.activeUser = aID;
        arsCwData.SelectUser();

        arsCwData.AssGrpTable.rows().every(

            function (rowIdx, tableLoop, rowLoop) {

                var rowData = this.data();

                if (rowData.Type == 1) {

                    baseGroups.push(rowData.aID)
                }
            }
        );

        if (baseGroups.length > 0) {

            $.ajax({

                data: { 'ba9b9529-3be4-4c03-ae33-e2deca05260a': baseGroups, 't': s }
            }).done(function (data) {
                
                arsCwData.lastResult = data;
                arsCwData.LoadTableData();
            });
        } else {

            arsBusy(true);
        }

    }

    arsCwData.RemoveBaseGroups = function (aID, s) {

        baseGroups = [];
        arsCwData.activeUser = aID;
        arsCwData.SelectUser();

        arsCwData.AssGrpTable.rows().every(

            function (rowIdx, tableLoop, rowLoop) {

                var rowData = this.data();

                if (rowData.Type == 1 && rowData.TTL > 0) {

                    baseGroups.push(rowData.aID)
                }
            }
        );

        if (baseGroups.length > 0) {

            arsBusy();

            $.ajax({

                data: { '642e2ac6-c179-44d0-9d56-1fe60abfc098': baseGroups, 't': s }
            }).done(function (data) {

                arsCwData.lastResult = data;
                arsCwData.LoadTableData();
                arsBusy(true);
            });
        }
    }

    arsCwData.RemoveAllGroups = function (aID, s) {

        allGroups = [];
        arsCwData.activeUser = aID;
        arsCwData.SelectUser();

        arsCwData.AssGrpTable.rows().every(

            function (rowIdx, tableLoop, rowLoop) {

                var rowData = this.data();

                if (rowData.TTL > 0) {

                    allGroups.push(rowData.aID)
                }
            }
        );

        if (allGroups.length > 0) {

            arsBusy();

            $.ajax({

                data: { '4d07fa2c-9eef-465d-930f-ea652aec8ca3': allGroups, 't': s }
            }).done(function (data) {

                arsCwData.lastResult = data;
                arsCwData.LoadTableData();
                arsBusy(true);
            });
        }
    }

    arsCwData.setGroupType = function (aID, f, n) {

        $.ajax({

            data: { '8c9219ab-deab-4076-8be7-55f4e7dd71a1': aID, 'f': f }
        }).done(function (data) {

            arsCwData.LoadTableData();
        })
    }

    arsCwData.AssUsrTable = $('#arsAssUsr').DataTable({

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

                    var s = $($('#arsTimeSelect').html());

                    $(td).append(
                        $('<div class="input-group"></div>').append(
                            $(arsCwData.Button).click(
                                function (event) {
                                    event.stopPropagation();
                                    arsCwData.RemoveAllGroups(rowData.aID, s.val());
                                }
                            ).text('Remove all')
                        ).append(
                            $(arsCwData.Button).click(
                                function (event) {
                                    event.stopPropagation();
                                    arsCwData.RemoveBaseGroups(rowData.aID, s.val());
                                }
                            ).text('Remove Base')
                        ).append(
                            $(arsCwData.Button).click(
                                function (event) {
                                    event.stopPropagation();
                                    arsCwData.RequestBaseGroups(rowData.aID, s.val());
                                }
                            ).text('Request Base')
                        ).append(s)
                    );
                },
                width: "1%"
            }
        ]
    }).on('click', 'tr', function (event) {

        var rowData = $(event.delegateTarget).DataTable().row(this).data();

        arsCwData.SelectUser(rowData.aID);
    });

    arsCwData.AssGrpTable = $('#arsAssGrp').DataTable({
        scrollY: '180px',
        scrollCollapse: true,
        paging: true,
        bLengthChange: false,
        deferRender: true,
        columnDefs: [
            {
                targets: 0,
                data: "Name",
                createdCell: function (td, cellData, rowData, row, col) {

                    $(td).text(cellData);
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
                        $(arsCwData.tplSwitch)
                    ).find('input').prop('checked', (rowData.Type == 1) ? true : false).on('click',
                        function (event) {
                            event.stopPropagation();
                            arsCwData.setGroupType(rowData.aID, (event.target.checked) ? 2 : 1, $(this));
                        }
                    );
                },
                width: "12%"
            },
            {
                targets: 3,
                data: "Expire",
                createdCell: function (td, cellData, rowData, row, col) {

                    if (cellData < 0) {

                        $(td).text("never");
                    } else if (cellData == 0) {

                        $(td).text("");
                    } else if (cellData > 0) {

                        $(td).text((new Date(rowData.Expire * 1000)).toLocaleString());
                    }
                },
                width: "1%"
            },
            {
                targets: 4,
                data: null,
                defaultContent: "",
                createdCell: function (td, cellData, rowData, row, col) {

                    if (rowData.TTL < 0) {

                        $(td).append(
                            $('<div class="input-group"></div>').append(
                                $(arsCwData.Button).click(
                                    function (event) {
                                        //event.stopPropagation();
                                        arsCwData.RemoveGroup(rowData.aID);
                                    }
                                ).text('Remove')
                            )
                        );
                    } else if (rowData.TTL == 0) {

                        $(td).append(
                            $('<div class="input-group"></div>').append(
                                $(arsCwData.Button).click(
                                    function (event) {
                                        //event.stopPropagation();
                                        arsCwData.RequestGroup(rowData.aID, $(this).siblings("select").first().val());
                                    }
                                ).text('Join')
                            ).append(arsCwData.TimeSelectHtml)
                        );
                    } else if (rowData.TTL > 0) {

                        $(td).append(
                            $('<div class="input-group"></div>').append(
                                $(arsCwData.Button).click(
                                    function (event) {
                                        //event.stopPropagation();
                                        arsCwData.RemoveGroup(rowData.aID);
                                    }
                                ).text('Remove')
                            ).append(
                                $(arsCwData.Button).click(
                                    function (event) {
                                        //event.stopPropagation();
                                        arsCwData.RequestGroup(rowData.aID, $(this).siblings("select").first().val());
                                    }
                                ).text('Update')
                            ).append(arsCwData.TimeSelectHtml)
                        );
                    }
                },
                width: "1%"

            }]
    });

    $('#arsActiveUserName').mouseenter(function (event) {
        $('#arsAssUsr').addClass('highlight-hover')
    }).mouseleave(function (event) {
        $('#arsAssUsr').removeClass('highlight-hover')
    });;

    $('.arsUserInfo button').click(function () {
        arsCwData.LoadTableData('380b841c-e14d-49f5-bdf2-0ee08d492215');
    });

    arsResize();
    setTimeout(arsCwData.LoadTableData, 50);
}

function arsAdvancedInit() {

    arsCwData = {};
    arsCwData.tableData = [];
    arsCwData.activeUser = null;
    arsCwData.Timeout = null

    arsCwData.UserName = $('#chUserName');
    arsCwData.UserInfo = $('#chUserInfo');

    arsCwData.PwBox = $('#tplArsPw').html();
    arsCwData.Button = $('#arsButton').html();
    arsCwData.AdvancedOptions = $('template[data-option]');
    
    arsResize = function () {

        $('.arsCsw .dataTables_scrollBody').css("max-height",
            $('#arsCw .arsManagedAccounts').first().height() - 160);

        $('#arsCwAdvancedOptions').css("max-height",
            $('.arsCsw .arsManagedAccounts').first().height() - 40);

        arsCwData.AssUsrTable.columns.adjust();
    }

    arsCwData.LoadTableData = function (guid = '0dc44518-4efc-4393-bb5c-85b4db04015f') {

        arsBusy();

        $.ajax({

            data: { [guid]: '' }
        }).done(function (data) {

            if ('Name' in data) {

                arsCwData.UserName.text(data.DisplayName);
                arsCwData.UserInfo.html('Department: <span>' + data.Department + '</span>, Company: <span>' + data.Company + '</span>');

                arsCwData.tableData = data.AssignedUsers;
                arsCwData.FillUserTable();
                arsCwData.SelectUser();

                arsBusy(true);
            } else {

                arsBusy(true);
            }
        });
    }

    arsCwData.FillUserTable = function () {

        var dt = arsCwData.AssUsrTable;

        dt.clear();
        dt.rows.add(arsCwData.tableData);
        dt.draw();
    }

    arsCwData.SelectUser = function (aID = null) {

        var entryFound = false;

        if (aID == null) {

            aID = arsCwData.activeUser;
        } else {

            arsCwData.activeUser = aID;
        }

        if (typeof aID != 'undefined' && aID != null) {

            arsCwData.AssUsrTable.rows().every(

                function (rowIdx, tableLoop, rowLoop) {

                    var rowData = this.data();

                    if (rowData.aID == aID) {

                        arsCwData.ShowAdvancedOptions(rowData);
                        entryFound = 1;
                    }
                }
            );
        }

        if (!entryFound) {

            arsCwData.activeUser = null;
            $('#arsCwAdvancedOptions').empty();
            $('#arsActiveUserName').text("<select Managed Account>");
        }
    }

    arsCwData.ShowAdvancedOptions = function (rowData) {

        var f = 0;
        var c = '';
        var a = $('#arsCwAdvancedOptions');

        a.html('');
        $('#arsActiveUserName').text("'" + rowData.DisplayName + "'");

        for (var i = 0; i < arsCwData.AdvancedOptions.length; i++) {

            var o = $(arsCwData.AdvancedOptions[i]).attr('data-option');

            if ((f % 2) == 0 && c != '') {

                a.append(
                    $('<div class="row"></div>').html(c)
                );
                c = '';
            }

            var n = $($(arsCwData.AdvancedOptions[i]).html());

            if ((rowData.flags & o) == o) {

                n.addClass('optDisabled');

                n.find('input, button').each(function () {

                   $(this).prop("disabled", true);
                });
            }

            f++;
            c += n[0].outerHTML;
        }

        if (c != '') {
            a.append(
                $('<div class="row"></div>').html(c)
            );
        }

        a.find("input,button").each(function () {

            var n = $(this);
            var o = n.attr('data-option');
            var i = n.attr('data-id');

            if (n[0].type == 'checkbox') {

                n.click(function () {

                    $.ajax({

                        data: { [i]: rowData.aID }
                    }).done(function (data) {

                        arsCwData.LoadTableData();
                    })
                })

                if ((rowData.flags & o) == o) {

                    n.prop('checked', true);
                }
            } else if (n[0].type == 'submit') {

                n.click(function () {

                    $.ajax({

                        data: { [i]: rowData.aID }
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
                            
                            p.find('input').val(data[arsCwData.activeUser]);

                            arsPopup('Reset password', p, [{ action: 'close', callback: function () { $('#arsModal').modal('hide'); }}]);
                        }
                    })
                })
            }
        })
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

    arsCwData.AssUsrTable = $('#arsAssUsr').DataTable({

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

                    if ( ( rowData.AccountStatus & 2 ) == 2 ) {

                        $(td).append('<i class="bi bi-file-excel" title="disabled"></i>');
                    } 

                    if ( ( rowData.AccountStatus & 16 ) == 16 ) {

                        $(td).append('<i class="bi bi-file-lock" title="locked"></i>');
                    } 

                    if ( ( rowData.AccountStatus & 1048576) == 1048576 ) {

                        $(td).append('<i class="bi bi-file-person-fill" title="no delegation"></i>');
                    }
                }
            },
            {
                targets: 3,
                data: null,
                defaultContent: "",
                createdCell: function (td, cellData, rowData, row, col) {

                    if (rowData.AccountStatus & 2) {

                        $(td).append(
                            $('<div class="input-group"></div>').append(
                                $(arsCwData.Button).click(
                                    function () {
                                        arsCwData.EnableAccount(rowData.aID);
                                    }
                                ).text('Enable Account')
                            )
                        );
                    } else {

                        $(td).append(
                            $('<div class="input-group"></div>').append(
                                $(arsCwData.Button).click(
                                    function () {
                                        arsCwData.DisableAccount(rowData.aID);
                                    }
                                ).text('Disable Account')
                            )
                        );
                    }
                    $(td).addClass("mw180");
                },
                width: "10%"
            }
        ]
    }).on('click', 'tr', function (event) {

        var rowData = $(event.delegateTarget).DataTable().row(this).data();

        arsCwData.activeUser = rowData.aID;
        arsCwData.ShowAdvancedOptions(rowData);
    });

    arsCwData.DisableAccount = function (aID) {

        $.ajax({

            data: { '218a74c3-c5ed-4e87-8a25-892dedd444dc': aID }
        }).done(function (data) {

            arsCwData.LoadTableData();
        })
    }

    arsCwData.EnableAccount = function (aID) {

        $.ajax({

            data: { 'd68b5138-626b-45be-8e31-9ed01ae20e37': aID }
        }).done(function (data) {

            arsCwData.LoadTableData();
        })
    }

    $('#arsActiveUserName').mouseenter(function (event) {
        $('#arsAssUsr').addClass('highlight-hover')
    }).mouseleave(function (event) {
        $('#arsAssUsr').removeClass('highlight-hover')
    });;

    $('.arsUserInfo button').click(function () {
        arsCwData.LoadTableData('ceb052c2-355e-40fe-9877-ff6a15705703');
    });

    arsResize();
    setTimeout(arsCwData.LoadTableData, 50);
}

function arsAuditInit() {

    arsCwData = {};

    arsResize = function () {
      
        arsCwData.AuditTable.columns.adjust();
    }

    arsCwData.AuditTable = $('#arsAuditLog').DataTable({
        processing: true,
        serverSide: true,
        ajax: {
            url: "./inc/asp/data.aspx",
            type: "POST",
            data: { '156d94b8-8bb3-41ac-8951-8d0ede66e075': '' }
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