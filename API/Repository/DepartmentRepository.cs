using API.Models;
using API.Repository.Interface;
using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace API.Repository
{
    public class DepartmentRepository : IDepartmentRepository
    {
        SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["BelajarAPI"].ConnectionString);

        DynamicParameters parameters = new DynamicParameters();
        public int Create(Department department) //insert to database
        {
            var procName = "SP_InsertDepartment";
            parameters.Add("@pDeptName", department.DepartmentName);
            var create = connection.Execute(procName, parameters, commandType: CommandType.StoredProcedure);
            return create;
        }


        public IEnumerable<Department> Get() //view database
        {
            var procName = "SP_ViewDepartment";
            var getalldept= connection.Query<Department>(procName, commandType: CommandType.StoredProcedure);
            return getalldept;
        }


        public async Task<IEnumerable<Department>> Get(int Id) //get id from database
        {
            var procName = "SP_GetDepartmentbyID";
            parameters.Add("@pDeptId", Id);

            var get = await connection.QueryAsync<Department>(procName, parameters, commandType: CommandType.StoredProcedure);
            return get;
        }

        public int Update(int Id, Department department) //update yo database
        {
            var procName = "SP_UpdateDepartment";
            parameters.Add("@pDeptId", Id);
            parameters.Add("@pDeptName", department.DepartmentName);
            var update = connection.Execute(procName, parameters, commandType: CommandType.StoredProcedure);
            return update;
        }

        public int Delete (int Id) //delete to database
        {
            var procName = "SP_DeleteDepartment";
            parameters.Add("@pDeptId", Id);
            var delete = connection.Execute(procName, parameters, commandType: CommandType.StoredProcedure);
            return delete;
        }

    }
}