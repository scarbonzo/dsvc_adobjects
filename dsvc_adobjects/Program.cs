using MongoDB.Driver;
using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Threading.Tasks;

class Program
{
    const int sleepTimeInMinutes = 2;

    static void Main(string[] args)
    {
        Console.WriteLine("Starting Service");

        while(true)
        {
            Console.WriteLine("Started run cycle @ " + DateTime.Now.ToString());
            //var des = GetCurrentADUserDEs(config.ActiveDirectoryUsername, config.ActiveDirectoryPassword);
            var adusers = GetCurrentADUsers(config.ActiveDirectoryUsername, config.ActiveDirectoryPassword);
            var dbusers = GetDatabaseUsers();

            //Process the existing DB users and update any mark deleted non-existent AD users
            try
            {
                Console.WriteLine("Started processing the DB users @ " + DateTime.Now.ToString());
                foreach (var dbuser in dbusers)
                {
                    if (dbuser.deleted != true)
                    {
                        var currentUser = adusers.FirstOrDefault(u => u.username == dbuser.username);
                        if (currentUser == null)
                        {
                            dbuser.deleted = true;
                            dbuser.deletiondate = DateTime.Now;
                        }
                    }
                }
                Console.WriteLine("Finished processing the DB users @ " + DateTime.Now.ToString());
            }
            catch (Exception e) { Console.WriteLine(e); }


            //Process all of the current AD users
            try
            {
                Console.WriteLine("Started updating DB @ " + DateTime.Now.ToString());
                //Process each of the ActiveDirectoryUser objects
                foreach (var aduser in adusers)
                {
                    try
                    {
                        //var user = CreateUserFromDE(aduser);
                        UpdateDatabaseUser(aduser);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
                Console.WriteLine("Finished updating DB @ " + DateTime.Now.ToString());
            }
            catch (Exception e) { Console.WriteLine(e); }

            Console.WriteLine("Finished run cycle @ " + DateTime.Now.ToString());
            Console.WriteLine("Sleeping for " + sleepTimeInMinutes.ToString() +" minute(s) @ " + DateTime.Now.ToString());
            Task.Delay(sleepTimeInMinutes * 60000).Wait();
        }
        //Console.ReadLine();
    }

    //Active Directory Functions
    //Returns a List<> of DirectoryEntries for all current AD users
    public static List<DirectoryEntry> GetCurrentADUserDEs(string Username, string Password)
    {
        
        var entries = new List<DirectoryEntry>();

        using (var context = new PrincipalContext(ContextType.Domain, "lsnj.org", Username, Password))
        {
            Console.WriteLine("Started scanning AD for users @ " + DateTime.Now.ToString());
            using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
            {
                var user = new UserPrincipal(context)
                {
                    Enabled = true //Only search for enabled accounts
                };
                searcher.QueryFilter = user;

                foreach (var result in searcher.FindAll())
                {
                    entries.Add(result.GetUnderlyingObject() as DirectoryEntry);
                }
            }
            Console.WriteLine("Finished scanning AD for users @ " + DateTime.Now.ToString());
        }

        return entries;
    }
    
    //Returns a List<> of DirectoryEntries for all current AD users
    public static List<ActiveDirectoryUser> GetCurrentADUsers(string Username, string Password)
    {
        var adusers = new List<ActiveDirectoryUser>();
        
        using (var context = new PrincipalContext(ContextType.Domain, "lsnj.org", Username, Password))
        {
            Console.WriteLine("Started scanning AD for users @ " + DateTime.Now.ToString());
            using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
            {
                var up = new UserPrincipal(context)
                {
                    Enabled = true //Only search for enabled accounts
                };
                searcher.QueryFilter = up;

                foreach (var result in searcher.FindAll())
                {
                    var user = CreateUserFromDE(result.GetUnderlyingObject() as DirectoryEntry);
                    adusers.Add(user);
                }
            }
            Console.WriteLine("Finished scanning AD for users @ " + DateTime.Now.ToString());
        }

        return adusers;
    }


    //Do some hocus pocus on strange Microsoft numbers
    private static Int64 GetInt64(DirectoryEntry de, string attr)
    {
        using (var ds = new DirectorySearcher(
            de,
            String.Format("({0}=*)", attr),
            new string[] { attr },
            SearchScope.Base
            ))
        {
            var sr = ds.FindOne();

            if (sr != null)
            {
                if (sr.Properties.Contains(attr))
                {
                    return (Int64)sr.Properties[attr][0];
                }
            }
            return -1;
        }
    }
        
    //Create an ActiveDirectoryUser object from a DirectoryEntry
    public static ActiveDirectoryUser CreateUserFromDE(DirectoryEntry de, MongoDB.Bson.ObjectId? Id = null)
    {
        var user = new ActiveDirectoryUser();

        try
        {
            if (Id != null)
            {
                user.Id = (MongoDB.Bson.ObjectId)Id;
            }
            else
            {
                user.Id = MongoDB.Bson.ObjectId.GenerateNewId();
            }
            try
            {
                user.username = de.Properties["samAccountName"].Value.ToString();
            }
            catch { }
            try
            {
                user.ou = de.Properties["distinguishedName"].Value.ToString();
            }
            catch { }
            try
            {
                user.displayname = de.Properties["displayName"].Value.ToString();
            }
            catch { }
            try
            {
                user.firstname = de.Properties["givenName"].Value.ToString();
            }
            catch { }
            try
            {
                user.lastname = de.Properties["sn"].Value.ToString();
            }
            catch { }
            try
            {
                user.passwordlastset = DateTime.FromFileTime(GetInt64(de, "pwdLastSet")); //If valid, update the Last Set Date
                if (user.passwordlastset < DateTime.Now.AddYears(-100))
                    user.passwordlastset = null;
            }
            catch { }
            try
            {
                user.lastlogin = DateTime.FromFileTime(GetInt64(de, "lastLogon")); //If valid, update the Last Login Date
                if (user.lastlogin < DateTime.Now.AddYears(-100))
                    user.lastlogin = null;
            }
            catch { }
            try
            {
                user.program = de.Properties["company"].Value.ToString();
            }
            catch { }
            try
            {
                user.office = de.Properties["physicalDeliveryOfficeName"].Value.ToString();
            }
            catch { }
            try
            {
                user.email = de.Properties["mail"].Value.ToString();
            }
            catch { }
            try
            {
                user.extension = de.Properties["telephoneNumber"].Value.ToString();
            }
            catch { }
            try
            {
                var memberGroups = de.Properties["memberOf"].Value;

                if (memberGroups.GetType() == typeof(string))
                {
                    user.groups = (string)memberGroups;
                }
                else if (memberGroups.GetType().IsArray)
                {
                    var memberGroupsEnumerable = memberGroups as IEnumerable;

                    if (memberGroupsEnumerable != null)
                    {
                        var asStringEnumerable = memberGroupsEnumerable.OfType<object>().Select(obj => obj.ToString());
                        user.groups = string.Join(", ", asStringEnumerable);
                    }
                }

                var _groups = new List<string>();

                foreach (string group in user.groups.Split(',').ToList().Where(g => g.Contains("CN=")).ToList())
                {
                    _groups.Add(group.Replace("CN=", "").Trim());
                }

                user.groups = string.Join(",", _groups);
                user.groupList = _groups.ToArray();
            }
            catch { }
            try
            {
                user.title = de.Properties["title"].Value.ToString();
            }
            catch { }
            try
            {
                if (Convert.ToBoolean((int)de.Properties["userAccountControl"].Value & 65536))
                {
                    user.expirable = false;
                }
                else
                {
                    user.expirable = true;
                }
            }
            catch { }
            try
            {
                user.notes = de.Properties["info"].Value.ToString();
            }
            catch { }

            user.active = true;
            user.lastupdate = DateTime.Now;
        }
        catch (Exception e) { Console.WriteLine(e); }

        return user;
    }

    //Database Functions
    //Write an ActiveDirectoryUser to MongoDB
    public static void UpdateDatabaseUser(ActiveDirectoryUser User)
    {
        //Collection hook
        var collection = new MongoClient(config.MongoDbConnectionString)
            .GetDatabase(config.MongoDbDatabase)
            .GetCollection<ActiveDirectoryUser>(config.MongoDbCollection);

        //Try to find an existing user to replace/update
        var existingUser = collection.AsQueryable()
            .FirstOrDefault(u => u.username == User.username);

        if(existingUser == null)
        {
            //Create a fresh user object
            collection.InsertOne(User);
        }
        else
        {
            //Replace an existing user object
            var filter = Builders<ActiveDirectoryUser>.Filter.Eq(s => s.Id, existingUser.Id);
            User.Id = existingUser.Id;

            collection.ReplaceOne(filter, User);
        }
    }

    //Get all DB Users
    public static List<ActiveDirectoryUser> GetDatabaseUsers()
    {
        var dbusers = new List<ActiveDirectoryUser>();

        //Collection hook
        var collection = new MongoClient(config.MongoDbConnectionString)
            .GetDatabase(config.MongoDbDatabase)
            .GetCollection<ActiveDirectoryUser>(config.MongoDbCollection)
            .AsQueryable();

        dbusers.AddRange(collection.ToList());
        
        return dbusers;
    }
}
