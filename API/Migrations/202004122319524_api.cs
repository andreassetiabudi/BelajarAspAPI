namespace API.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class api : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.TB_M_Department", "UpdateDate", c => c.DateTimeOffset(nullable: false, precision: 7));
            AlterColumn("dbo.TB_M_Department", "DeleteDate", c => c.DateTimeOffset(nullable: false, precision: 7));
            AlterColumn("dbo.TB_M_Division", "UpdateDate", c => c.DateTimeOffset(nullable: false, precision: 7));
            AlterColumn("dbo.TB_M_Division", "DeleteDate", c => c.DateTimeOffset(nullable: false, precision: 7));
        }
        
        public override void Down()
        {
            AlterColumn("dbo.TB_M_Division", "DeleteDate", c => c.DateTimeOffset(precision: 7));
            AlterColumn("dbo.TB_M_Division", "UpdateDate", c => c.DateTimeOffset(precision: 7));
            AlterColumn("dbo.TB_M_Department", "DeleteDate", c => c.DateTimeOffset(precision: 7));
            AlterColumn("dbo.TB_M_Department", "UpdateDate", c => c.DateTimeOffset(precision: 7));
        }
    }
}
