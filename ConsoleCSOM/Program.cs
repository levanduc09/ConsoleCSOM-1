using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM
{
    class SharepointInfo
    {
        public string SiteUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = GetContext(clientContextHelper);
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();

                    // create a list called "CSOM list"
                    await CreateListAsync(ctx, "CSOM list", ListTemplateType.GenericList);
                    // create a term set
                    await CreateTermSetAsync(ctx);
                    // create site content type with id = "0x0100A33D9AD9805788419BDAAC2CCB37509E"
                    await CreateSiteContentTypeAsync(ctx, "0x0100A33D9AD9805788419BDAAC2CCB37509E");
                    // create site fields "about: text" and "city: taxonomy" and add it to the content type above
                    await CreateSiteFieldAsync(ctx);
                    // apply content type to the list
                    await ApplyContentTypeAsync(ctx, "CSOM list", "0x0100A33D9AD9805788419BDAAC2CCB37509E");
                    // add 5 items to the list "CSOM list"
                    await AddListItemsAsync(ctx, "CSOM list");
                    // update site fields
                    await UpdateSiteFieldsAsync(ctx);
                    // CAML query
                    await SimpleCamlQueryAsync(ctx, "CSOM list");
                    // create list view
                    await CreateListViewByCSOMAsync(ctx, "CSOM list");
                    // bulk update
                    await UpdateBatchDataAsync(ctx, "CSOM list");
                    // create taxonomy field with multiple values
                    await CreateTaxonomyFieldWithMultipleValuesAsync(ctx);
                    // add 3 items to the list "CSOM list" and set multiple values for "Cities"
                    await AddListItemsAsyncAndSetValuesForCities(ctx, "CSOM list");
                    // create new list with type = document
                    await CreateListAsync(ctx, "Document Test", ListTemplateType.DocumentLibrary);
                    // apply content type to the list above
                    await ApplyContentTypeAsync(ctx, "Document Test", "0x0100A33D9AD9805788419BDAAC2CCB37509E");
                    // create folder
                    await CreateFolderAsync(ctx, "Document Test");
                    // create site field "CSOMAuthor"
                    await AddCSOMAuthorAsync(ctx);

                }

                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
            }
        }

        public static async Task<File> UploadFile(ClientContext ctx, Folder folder)
        {
            ctx.Load(folder, f => f.ServerRelativeUrl);
            await ctx.ExecuteQueryAsync();
            var fileCreationInfo = new FileCreationInformation
            {
                Content = Encoding.ASCII.GetBytes("test"),
                Overwrite = true,
                Url = folder.ServerRelativeUrl + "/test.txt"
            };

            var newFile = folder.Files.Add(fileCreationInfo);
            await ctx.ExecuteQueryAsync();
            return newFile;
        }

        private static async Task CreateFolderAsync(ClientContext context, string listTitle)
        {
            File file;
            List list = context.Web.Lists.GetByTitle(listTitle);

            FolderCollection folders = list.RootFolder.Folders;
            Folder folder1 = folders.Add("Folder 1");
            Folder folder2 = folder1.Folders.Add("Folder 2");

            // upload files to folder1
            file = await UploadFile(context, folder1);
            file.ListItemAllFields["About"] = "Folder test";
            // Stockholm
            file.ListItemAllFields["City"] = $"0679a780-032a-4e5b-b5f5-95762a8b082c";
            file.ListItemAllFields.Update();

            // upload files to folder2
            file = await UploadFile(context, folder2);
            file.ListItemAllFields["About"] = "Folder test";
            // Ho Chi Minh
            file.ListItemAllFields["Cities"] = $"d2ede839-c7df-4951-9eb6-6fc24f02cb44";
            file.ListItemAllFields.Update();
            await context.ExecuteQueryAsync();

            // get items in folder2
            // todo
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"
                                    <View>
                                        <Query>
                                            <Where>
                                                <Eq>
                                                    <FieldRef Name='Cities' />
                                                    <Value Type='Text'>Stockholm</Value>
                                                </Eq>
                                            </Where>
                                        </Query>
                                    </View>
                                ";
            camlQuery.FolderServerRelativeUrl = folder2.ServerRelativeUrl;
            var items = list.GetItems(camlQuery);
            context.Load(items);
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateSiteFieldAsync(ClientContext context)
        {
            #region[Variables]
            string[] trgtfld = { "About", "City" };
            FieldCollection fields = context.Web.Fields;
            //context.Load(fields);
            FieldLinkCreationInformation fldLink = null;
            Field txtFldAbout;
            Field taxonomyFieldCity;
            string taxonomyFieldCityAsXML = string.Empty;
            string txtFieldAboutAsXML = string.Empty;
            ContentType contentType = context.Web.ContentTypes.GetById("0x0100A33D9AD9805788419BDAAC2CCB37509E");
            #endregion

            // Add site fields "About", "City"
            // check if "About" exists
            context.Load(fields, fields => fields.Include(f => f.InternalName).Where(f => f.InternalName == "About"));
            await context.ExecuteQueryAsync();
            if (fields.Count == 0)
            {
                txtFieldAboutAsXML = @"<Field Name='About' DisplayName='About' Type='Text' Hidden='False' Group='CSOM' />";
                txtFldAbout = fields.AddFieldAsXml(txtFieldAboutAsXML, true, AddFieldOptions.DefaultValue);
            }
            // check if City exists
            fields = context.Web.Fields;
            context.Load(fields, fields => fields.Include(f => f.InternalName).Where(f => f.InternalName == "City"));
            await context.ExecuteQueryAsync();
            if (fields.Count == 0)
            {
                taxonomyFieldCityAsXML = "<Field Type='TaxonomyFieldType' Name='City' StaticName='City' DisplayName = 'City' />";
                taxonomyFieldCity = fields.AddFieldAsXml(taxonomyFieldCityAsXML, true, AddFieldOptions.AddFieldInternalNameHint);
                // Retrieve the field as a Taxonomy Field
                TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(taxonomyFieldCity);
                taxonomyField.SspId = new Guid("29c02e79cb814fddb38201d3046895a2");
                taxonomyField.TermSetId = new Guid("5c1cddeb-ce15-4604-96d2-54e6c6e553b4");
                taxonomyField.TargetTemplate = String.Empty;
                taxonomyField.AnchorId = Guid.Empty;
                taxonomyField.Update();
            }
            await context.ExecuteQueryAsync();
            foreach (var fld in trgtfld)
            {
                try
                {
                    Field field = context.Web.AvailableFields.GetByInternalNameOrTitle(fld);
                    context.Load(field, f => f.Id);
                    await context.ExecuteQueryAsync();
                    if (contentType.FieldLinks.GetById(field.Id) == null)
                    {
                        fldLink = new FieldLinkCreationInformation();
                        fldLink.Field = field;
                        contentType.FieldLinks.Add(fldLink);
                        contentType.Update(true);
                        await context.ExecuteQueryAsync();
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }

            }
        }

        private static async Task CreateTaxonomyFieldWithMultipleValuesAsync(ClientContext context)
        {
            #region[Variables]
            string[] trgtfld = { "Cities" };
            FieldCollection fields = context.Web.Fields;
            FieldLinkCreationInformation fldLink = null;
            Field taxonomyFieldCities;
            string taxonomyFieldCitiesAsXML = string.Empty;
            ContentType contentType = context.Web.ContentTypes.GetById("0x0100A33D9AD9805788419BDAAC2CCB37509E");
            #endregion

            context.Load(fields, fields => fields.Include(f => f.InternalName).Where(f => f.InternalName == "Cities"));
            await context.ExecuteQueryAsync();
            // Add site field "Cities"
            if (fields.Count == 0)
            {
                taxonomyFieldCitiesAsXML = "<Field Type='TaxonomyFieldType' Name='Cities' StaticName='Cities' DisplayName = 'Cities' />";
                taxonomyFieldCities = fields.AddFieldAsXml(taxonomyFieldCitiesAsXML, true, AddFieldOptions.AddFieldInternalNameHint);
                // Retrieve the field as a Taxonomy Field
                TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(taxonomyFieldCities);
                taxonomyField.SspId = new Guid("29c02e79cb814fddb38201d3046895a2");
                taxonomyField.TermSetId = new Guid("5c1cddeb-ce15-4604-96d2-54e6c6e553b4");
                taxonomyField.TargetTemplate = String.Empty;
                taxonomyField.AnchorId = Guid.Empty;
                taxonomyField.AllowMultipleValues = true;
                taxonomyField.Update();
                await context.ExecuteQueryAsync();
            }
            else
            {
                taxonomyFieldCities = fields.FirstOrDefault();
            }
            foreach (var fld in trgtfld)
            {
                try
                {
                    FieldLinkCollection fieldLinks = contentType.FieldLinks;
                    context.Load(fieldLinks, f => f.Where(f => f.Id == taxonomyFieldCities.Id));
                    await context.ExecuteQueryAsync();
                    if (fieldLinks.Count == 0)
                    {
                        fldLink = new FieldLinkCreationInformation
                        {
                            Field = taxonomyFieldCities
                        };
                        contentType.FieldLinks.Add(fldLink);
                        contentType.Update(true);
                        await context.ExecuteQueryAsync();
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }

            }
        }

        private static async Task UpdateSiteFieldsAsync(ClientContext context)
        {
            #region[Variables]
            FieldCollection fields = context.Web.Fields;
            context.Load(fields);
            Field txtFldAbout = fields.GetByInternalNameOrTitle("About");
            context.Load(txtFldAbout, f => f.TypeAsString);
            Field taxonomyFieldCity = fields.GetByInternalNameOrTitle("City");
            context.Load(taxonomyFieldCity, t => t.TypeAsString);
            await context.ExecuteQueryAsync();
            #endregion

            // set default
            if (txtFldAbout != null)
            {
                txtFldAbout.DefaultValue = "About Default";
                txtFldAbout.Update();
                txtFldAbout.UpdateAndPushChanges(true);
            }
            if (taxonomyFieldCity != null)
            {
                // Retrieve the field as a Taxonomy Field
                var taxColumn = context.CastTo<TaxonomyField>(taxonomyFieldCity);
                context.Load(taxColumn);
                await context.ExecuteQueryAsync();
                //initialize taxonomy field value
                var defaultValue = new TaxonomyFieldValue();
                defaultValue.WssId = -1;
                defaultValue.Label = "Ho Chi Minh";
                defaultValue.TermGuid = "d2ede839-c7df-4951-9eb6-6fc24f02cb44";
                //retrieve validated taxonomy field value
                var validatedValue = taxColumn.GetValidatedString(defaultValue);
                await context.ExecuteQueryAsync();
                //set default value for a taxonomy field
                taxColumn.DefaultValue = validatedValue.Value;
                taxColumn.Update();
                taxColumn.UpdateAndPushChanges(true);
                await context.ExecuteQueryAsync();
            }
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateListAsync(ClientContext context, string listTitle, ListTemplateType listTemplateType)
        {
            ListCollection listCollection = context.Web.Lists;
            context.Load(listCollection, lists => lists.Include(list => list.Title).Where(list => list.Title == listTitle));
            await context.ExecuteQueryAsync();
            if (listCollection.Count == 0)
            {
                Web web = context.Web;
                ListCreationInformation listCreationInformation = new ListCreationInformation();
                listCreationInformation.Title = listTitle;
                listCreationInformation.TemplateType = (int)listTemplateType;
                List list = web.Lists.Add(listCreationInformation);
                await context.ExecuteQueryAsync();
            }
        }

        private static async Task AddListItemsAsync(ClientContext context, string listTitle)
        {
            ListCollection listCollection = context.Web.Lists;
            context.Load(listCollection, lists => lists.Include(list => list.Title).Where(list => list.Title == listTitle));
            await context.ExecuteQueryAsync();
            if (listCollection.Count > 0)
            {
                List list = listCollection.FirstOrDefault();
                // add 5 items to the list above
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem;
                for (int i = 1; i <= 5; i++)
                {
                    oListItem = list.AddItem(itemCreateInfo);
                    oListItem["Title"] = $"Item{i}";
                    oListItem["About"] = $"Item{i}";
                    oListItem["City"] = $"d2ede839-c7df-4951-9eb6-6fc24f02cb44";
                    oListItem["ContentTypeId"] = "0x0100A33D9AD9805788419BDAAC2CCB37509E";
                    oListItem.Update();
                }
                await context.ExecuteQueryAsync();
            }
        }

        private static async Task AddListItemsAsyncAndSetValuesForCities(ClientContext context, string listTitle)
        {
            ListCollection listCollection = context.Web.Lists;
            context.Load(listCollection, lists => lists.Include(list => list.Title).Where(list => list.Title == listTitle));
            await context.ExecuteQueryAsync();
            if (listCollection.Count > 0)
            {
                List list = listCollection.FirstOrDefault();
                // add 3 items to the list above
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem;
                for (int i = 1; i <= 3; i++)
                {
                    oListItem = list.AddItem(itemCreateInfo);
                    oListItem["Title"] = $"Item{i}";
                    oListItem["About"] = $"Item{i}";
                    oListItem["City"] = "d2ede839-c7df-4951-9eb6-6fc24f02cb44";
                    oListItem["Cities"] = "-1;#Ho Chi Minh|d2ede839-c7df-4951-9eb6-6fc24f02cb44;#-1;#Stockholm|0679a780-032a-4e5b-b5f5-95762a8b082c";
                    oListItem["ContentTypeId"] = "0x0100A33D9AD9805788419BDAAC2CCB37509E";
                    oListItem.Update();
                }
                await context.ExecuteQueryAsync();
            }
        }

        private static async Task CreateSiteContentTypeAsync(ClientContext ctx, string contentTypeId)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            ctx.Load(contentTypes);
            await ctx.ExecuteQueryAsync();

            foreach (var item in contentTypes)
            {
                if (item.StringId == contentTypeId)
                    return;
            }

            // Create a Content Type Information object.
            ContentTypeCreationInformation newCt = new ContentTypeCreationInformation();
            // Set the name for the content type.
            newCt.Name = "CSOM test content type";
            // Set id for the content type.
            newCt.Id = contentTypeId;
            // Set content type to be available from specific group.
            newCt.Group = "CSOM test content type group";
            // Create the content type.
            ContentType myContentType = contentTypes.Add(newCt);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task ApplyContentTypeAsync(ClientContext context, string listName, string contentTypeId)
        {
            #region[Variables]
            ListCollection spListColl = null;
            ContentType spContentType = null;
            #endregion
            try
            {
                spListColl = context.Web.Lists;
                context.Load(spListColl);
                await context.ExecuteQueryAsync();
                foreach (List list in spListColl)
                {
                    try
                    {
                        if (list.Title == listName)
                        {
                            list.ContentTypesEnabled = true;
                            // to do
                            if (list.ContentTypes.GetById(contentTypeId) == null)
                            {
                                spContentType = context.Web.ContentTypes.GetById(contentTypeId);
                                list.ContentTypes.AddExistingContentType(spContentType);
                                list.Update();
                                context.Web.Update();
                                await context.ExecuteQueryAsync();
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

        private static async Task GetFieldTermValue(ClientContext Ctx, string termId)
        {
            //load term by id
            TaxonomySession session = TaxonomySession.GetTaxonomySession(Ctx);
            Term taxonomyTerm = session.GetTerm(new Guid(termId));
            Ctx.Load(taxonomyTerm, t => t.Labels,
                                   t => t.Name,
                                   t => t.Id);
            await Ctx.ExecuteQueryAsync();
        }

        private static async Task ExampleSetTaxonomyFieldValue(ListItem item, ClientContext ctx)
        {
            var field = ctx.Web.Fields.GetByTitle("fieldname");

            ctx.Load(field);
            await ctx.ExecuteQueryAsync();

            var taxField = ctx.CastTo<TaxonomyField>(field);

            taxField.SetFieldValueByValue(item, new TaxonomyFieldValue()
            {
                WssId = -1, // alway let it -1
                Label = "correct label here",
                TermGuid = "term id"
            });
            item.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CreateTermSetAsync(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("Test Term Set");
            // Create a new term set called city-{your name}
            TermSet myTermSet;
            if (termGroup.TermSets.GetByName("city-dule") == null)
            {
                myTermSet = termGroup.CreateTermSet("city-dule", Guid.NewGuid(), 1033);
            }
            else
            {
                myTermSet = termGroup.TermSets.GetByName("city-dule");
            }
            // add 2 terms "Ho Chi Minh" and "Stockholm" to the term set above
            if (myTermSet.Terms.GetByName("Stockholm") == null)
            {
                myTermSet.CreateTerm("Stockholm", 1033, Guid.NewGuid());
            }
            if (myTermSet.Terms.GetByName("Ho Chi Minh") == null)
            {
                myTermSet.CreateTerm("Ho Chi Minh", 1033, Guid.NewGuid());
            }
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomLinqAsync(ClientContext ctx)
        {
            var fieldsQuery = from f in ctx.Web.Fields
                              where f.InternalName == "Test" ||
                                    f.TypeAsString == "TaxonomyFieldTypeMulti" ||
                                    f.TypeAsString == "TaxonomyFieldType"
                              select f;

            var fields = ctx.LoadQuery(fieldsQuery);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task SimpleCamlQueryAsync(ClientContext ctx, string listTitle)
        {
            var list = ctx.Web.Lists.GetByTitle(listTitle);

            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <Where>
                                        <Neq>
                                            <FieldRef Name='About' />
                                            <Value Type='Text'>About Default</Value>
                                        </Neq>
                                    </Where>
                                </Query>
                            </View>"
            });

            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task AddCSOMAuthorAsync(ClientContext context)
        {
            var list = context.Web.Lists.GetByTitle("CSOM list");

            FieldCollection fields = list.Fields;
            context.Load(fields, fields => fields.Where(f => f.InternalName == "CSOMAuthor"));
            await context.ExecuteQueryAsync();

            // check if field exists
            if (fields.Count == 0)
            {
                string field = @"<Field Name='CSOMAuthor' DisplayName='CSOMAuthor' Type='User' Group='CSOM' />";
                list.Fields.AddFieldAsXml(field, true, AddFieldOptions.DefaultValue);
            }

            // Update data
            var allItemQuery = CamlQuery.CreateAllItemsQuery();
            var items = list.GetItems(allItemQuery);
            context.Load(items);
            await context.ExecuteQueryAsync();

            var user = context.Web.EnsureUser("admin@omniapreprod.onmicrosoft.com");

            foreach (var item in items)
            {
                item["CSOMAuthor"] = user;
                item.Update();
            }
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateListViewByCSOMAsync(ClientContext context, string listTitle)
        {
            string viewQuery = @"<Where>
                                    <Eq>
                                        <FieldRef Name='City' />
                                        <Value Type='Text'>Ho Chi Minh</Value>
                                    </Eq>
                                </Where>
                                <OrderBy>
                                    <FieldRef Name='Modified' Ascending='False' />
                                </OrderBy> ";
            var list = context.Web.Lists.GetByTitle(listTitle);
            ViewCollection views = list.Views;
            if (views.GetByTitle("CSOM View") != null)
            {

                ViewCreationInformation creationInfo = new ViewCreationInformation
                {
                    Title = "CSOM View",
                    RowLimit = 50,
                    ViewFields = new string[] { "ID", "Title", "City", "About" },
                    ViewTypeKind = ViewType.None,
                    SetAsDefaultView = true,
                    Query = viewQuery
                };
                views.Add(creationInfo);
                await context.ExecuteQueryAsync();
            }
        }

        private static async Task UpdateBatchDataAsync(ClientContext context, string listTitle)
        {
            var list = context.Web.Lists.GetByTitle(listTitle);
            CamlQuery camlQuery = new()
            {
                ViewXml = @"<View>
                                <Query>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name='About' />
                                            <Value Type='Text'>About Default</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                            </View>"
            };

            var items = list.GetItems(camlQuery);
            context.Load(items);
            await context.ExecuteQueryAsync();

            for (var i = 0; i < items.Count; i++)
            {
                items[i]["About"] = "Update Script";
                items[i].Update();
                if ((i + 1) % 2 == 0 || (i == items.Count - 1))
                {
                    await context.ExecuteQueryAsync();
                }
            }
        }
    }
}
