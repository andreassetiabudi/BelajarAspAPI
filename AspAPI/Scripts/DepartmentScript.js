$(document).ready(function () {
    $('#Department').dataTable({
        "ajax": {
            url: "/Department/LoadDepartment",
            type: "GET",
            dataType: "json",
            dataSrc: "",
        },
        "columns": [
            { "name": "Department Name", "data": "DepartmentName"},
            { "data": "CreateDate", "render": function (data) { return moment(data).format('MMMM Do YYYY, h:mm:ss a');}},
            { "data": "UpdateDate", "render": function (data) 
            { var dateupdate = "Not Updated Yet";
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
                url: "/Department/Delete/",
                type: "POST",
                data: { id: Id }
            }).then((result) => {
                if (result.StatusCode == 200) {
                    Swal.fire({
                        icon: 'success',
                        position: 'center',
                        title: 'Department Delete Successfully',
                    }).then((result) => {
                        if (result.value) {
                            $('#Department').DataTable().ajax.reload();
                        }
                    });
                }
                else {
                    Swal.fire({
                        icon: 'error',
                        title: 'error',
                        text: 'Failed to Delete Department ',
                    })
                    ClearScreen();
                }
            })
        }
    });
}

function Save() {
    if ($('#DepartmentName').val() == 0) {
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'Something went wrong!',
            footer: '<a href>Please Full Fill The Department Name</a>'
        });
        $('#Department').DataTable().ajax.reload();
    }
    else {
        var Department = new Object();
        Department.Id = $('#Id').val();
        Department.DepartmentName = $('#DepartmentName').val();
        $.ajax({
            url: "/Department/InsertOrUpdate/",
            type: "POST",
            data: Department
        }).then((result) => {
            //debugger;
            if (result.StatusCode == 200) {
                Swal.fire({
                    position: 'center',
                    type: 'success',
                    title: 'Department Added Successfully',
                }).then((result) => {
                    if (result.value) {
                        $('#Department').DataTable().ajax.reload();
                    }
                });
            }
            else {
                Swal.fire('Error', 'Failed to Add Department', 'error');
                ClearScreen();
            }
        });
    }
}





function ClearScreen() {
    $('#Id').val('');
    $('#DepartmentName').val('');
    $('#EditBtn').hide();
    $('#SaveBtn').show();
}




function GetById(Id) {
    $.ajax({
        url: "/Department/GetById/" + Id,
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        dataType: "json",
        async: false,
        success: function (result) {
            const obj = JSON.parse(result);
            $('#Id').val(obj.Id);
            $('#DepartmentName').val(obj.Name);
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
    if ($('#DepartmentName').val() == 0) {
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'Something went wrong!',
            footer: '<a href>Please Full Fill The Department Name</a>'
        });
        $('#Department').DataTable().ajax.reload();
    }
    else {
        var Department = new Object();
        Department.Id = $('#Id').val();
        Department.DepartmentName = $('#DepartmentName').val();
        $.ajax({
            url: "/Department/InsertOrUpdate/",
            type: "POST",
            data: Department
        }).then((result) => {
            if (result.StatusCode == 200) {
                Swal.fire({
                    position: 'center',
                    type: 'success',
                    title: 'Department Updated Successfully',
                }).then((result) => {
                    if (result.value) {
                        $('#Department').DataTable().ajax.reload();
                    }
                });
            }
            else {
                Swal.fire('Error', 'Failed to Update Department', 'error');
                $('#myModal').modal('show');
                $('#SaveBtn').hide();
            }
        });
    }
}
