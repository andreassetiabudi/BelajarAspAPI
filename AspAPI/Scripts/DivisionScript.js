$(document).ready(function () {
    $('#Division').dataTable({
        "ajax": {
            url: "/Division/LoadDivision",
            type: "GET",
            dataType: "json",
            dataSrc: "",
        },
        "columns": [
            { "name": "Division Name", "data": "DivisionName"},
            { "name": "Department Name", "data": "DepartmentName"},
            { "data": "CreateDate", "render": function (data) { return moment(data).format('MMMM Do YYYY, h:mm:ss a'); } },
            {
                "data": "UpdateDate", "render": function (data) {
                    var dateupdate = "Not Updated Yet";
                    var nulldate = null;
                    if (data == nulldate) {
                        return dateupdate;
                    } else {
                        return moment(data).format('MMMM Do YYYY, h:mm:ss a');
                    }
                }
            },
            {
                data: null, render: function (data, type, dept) {
                    return " <td><button type='button' class='btn btn-warning' id='Update' onclick=$('#EditBtn').show();GetById('" + dept.Id + "');>Edit</button> <button type='button' class='btn btn-danger' id='Delete' onclick=Delete('" + dept.Id + "');>Delete</button ></td >";
                }
            },
        ]
    });
});

function Delete(Id) {
    Swal.fire({
        title: "Do you want to delete it?",
        text: "You won't be able to restore this!",
        showCancelButton: true,
        confirmButtonColor: '#3085d6',
        cancelButtonColor: '#d33',
        confirmButtonText: 'Yes, delete it!'
    }).then((result) => {
        if (result.value) {
            $.ajax({
                url: "/Division/Delete/",
                type: "POST",
                data: { id: Id }
            }).then((result) => {
                if (result.StatusCode == 200) {
                    Swal.fire({
                        icon: 'success',
                        position: 'center',
                        title: 'Division Delete Successfully',
                    }).then((result) => {
                        if (result.value) {
                            $('#Division').DataTable().ajax.reload();
                        }
                    });
                }
                else {
                    Swal.fire({
                        icon: 'error',
                        title: 'error',
                        text: 'Failed to Delete Division',
                    })
                    ClearScreen();
                }
            })
        }
    });
}

function Save() {
    if ($('#DivisionName').val() == 0) {
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'Something went wrong!',
            footer: '<a href>Please Full Fill The Division Name</a>'
        });
        $('#Division').DataTable().ajax.reload();
    }
    else {


        var Division = new Object();
        Division.Id = $('#Id').val();
        Division.DivisionName = $('#DivisionName').val();
        Division.Department_Id = $('#Department_Id').val();
        $.ajax({
            url: "/Division/InsertOrUpdate/",
            type: "POST",
            data: Division
        }).then((result) => {
            //debugger;
            if (result.StatusCode == 200) {
                Swal.fire({
                    position: 'center',
                    type: 'success',
                    title: 'Division Added Successfully',
                }).then((result) => {
                    if (result.value) {
                        $('#Division').DataTable().ajax.reload();
                    }
                });
            }
            else {
                Swal.fire('Error', 'Failed to Add Division', 'error');
                ClearScreen();
            }
        });
    }
}





function ClearScreen() {
    $('#Id').val('');
    $('#DivisionName').val('');
    $('#EditBtn').hide();
    $('#SaveBtn').show();
}




function GetById(Id) {
    $.ajax({
        url: "/Division/GetById/" + Id,
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        dataType: "json",
        async: false,
        success: function (result) {
            const obj = JSON.parse(result);
            $('#Id').val(obj.Id);
            $('#DivisionName').val(obj.Name);
            $('#Department_Id').val(obj.Department_Id);
            $('#myModal').modal('show');
            $('#UpdateBtn').show();
            $('#SaveBtn').hide();
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
    return false;
}


function Edit() {
    if ($('#DivisionName').val() == 0) {
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'Something went wrong!',
            footer: '<a href>Please Full Fill The Division Name</a>'
        });
        $('#Division').DataTable().ajax.reload();
    }
    else {
        var Division = new Object();
        Division.Id = $('#Id').val();
        Division.DivisionName = $('#DivisionName').val();
        Division.Department_Id = $('#Department_Id').val();
        $.ajax({
            url: "/Division/InsertOrUpdate/",
            type: "POST",
            data: Division
        }).then((result) => {
            if (result.StatusCode == 200) {
                Swal.fire({
                    position: 'center',
                    type: 'success',
                    title: 'Division Updated Successfully',
                }).then((result) => {
                    if (result.value) {
                        $('#Division').DataTable().ajax.reload();
                    }
                });
            }
            else {
                Swal.fire('Error', 'Failed to Update Division', 'error');
                $('#myModal').modal('show');
                $('#SaveBtn').hide();
            }
        });
    }
}


var Departments = []
function LoadDepartment(element) {
    if (Departments.length == 0) {
        $.ajax({
            type: "Get",
            url: "/Department/LoadDepartment",
            success: function (data) {
                Departments = data;
                renderDepartment(element);
            }
        })
    }
    else {
        renderDepartment(element);
    }
}

function renderDepartment(element) {
    var $ele = $(element);  
    $ele.empty();
    $ele.append($('<option/>').val('0').text('Select Department').hide());
    $.each(Departments, function (i, val) {
        $ele.append($('<option/>').val(val.Id).text(val.DepartmentName));
    })
}

LoadDepartment($('#Department_Id'));

