using System;
using System.Collections.Generic;
using System.Text;

public class ActiveDirectoryUser
{
    public MongoDB.Bson.ObjectId Id { get; set; }
    public string username { get; set; }
    public string displayname { get; set; }
    public string firstname { get; set; }
    public string lastname { get; set; }
    public string ou { get; set; }
    public string program { get; set; }
    public string office { get; set; }
    public string email { get; set; }
    public string extension { get; set; }
    public string title { get; set; }
    public string groups { get; set; }
    public string[] groupList { get; set; }
    public Nullable<System.DateTime> passwordlastset { get; set; }
    public Nullable<System.DateTime> lastlogin { get; set; }
    public Nullable<bool> deleted { get; set; }
    public Nullable<System.DateTime> deletiondate { get; set; }
    public string notes { get; set; }
    public Nullable<bool> expirable { get; set; }
    public Nullable<System.DateTime> lastupdate { get; set; }
    public Nullable<bool> active { get; set; }
}
