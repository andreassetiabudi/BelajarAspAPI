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
using API.ViewModel;
using System.Web;

namespace API.Repository
{
    public class DivisionRepository : IDivisionRepository
    {
        SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["BelajarAPI"].ConnectionString);

        DynamicParameters parameters = new DynamicParameters();
        public int Create(Division division)
        {
            var procName = "SP_InsertDivision";
            parameters.Add("@pDivName", division.DivisionName);
            parameters.Add("@pDeptId", division.Department_Id);
            var create = connection.Execute(procName, parameters, commandType: CommandType.StoredProcedure);
            return create;
        }

        public IEnumerable<DivisionVM> Get()
        {
            var procName = "SP_ViewDivision";
            var getalldiv = connection.Query<DivisionVM>(procName, commandType: CommandType.StoredProcedure);
            return getalldiv;
        }


        public async Task<IEnumerable<DivisionVM>> Get(int Id)
        {
            var procName = "SP_GetDivisionbyID";
            parameters.Add("@pDivId", Id);

            var get = await connection.QueryAsync<DivisionVM>(procName, parameters, commandType: CommandType.StoredProcedure);
            return get;
        }

        public int Update(int Id, Division division)
        {
            var procName = "SP_UpdateDivision";
            parameters.Add("@pDivId", Id);
            parameters.Add("@pDivName", division.DivisionName);
            parameters.Add("@pDeptId", division.Department_Id);
            var update = connection.Execute(procName, parameters, commandType: CommandType.StoredProcedure);
            return update;
        }

        public int Delete(int Id)
        {
            var procName = "SP_DeleteDivision";
            parameters.Add("@pDivId", Id);
            var delete = connection.Execute(procName, parameters, commandType: CommandType.StoredProcedure);
            return delete;
        }
    }
}