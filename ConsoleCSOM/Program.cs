using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Linq;
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

                    // create a list
                    await CreateListAsync(ctx, "CSOM list");
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
                    //await SimpleCamlQueryAsync(ctx);
                }

                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
            }
        }

        private static async Task CreateSiteFieldAsync(ClientContext context)
        {
            #region[Variables]
            string[] trgtfld = { "About", "City" };
            FieldCollection fields = context.Web.Fields;
            context.Load(fields);
            FieldLinkCreationInformation fldLink = null;
            Field txtFldAbout = fields.GetByInternalNameOrTitle("About");
            context.Load(txtFldAbout, t => t.TypeAsString);
            Field taxonomyFieldCity = fields.GetByInternalNameOrTitle("City");
            string taxonomyFieldCityAsXML = string.Empty;
            string txtFieldAboutAsXML = string.Empty;
            ContentType contentType = context.Web.ContentTypes.GetById("0x0100A33D9AD9805788419BDAAC2CCB37509E");
            #endregion

            // Add site fields "About", "City"
            if (txtFldAbout == null)
            {
                txtFieldAboutAsXML = @"<Field Name='About' DisplayName='About' Type='Text' Hidden='False' Group='CSOM' />";
                txtFldAbout = fields.AddFieldAsXml(txtFieldAboutAsXML, true, AddFieldOptions.DefaultValue);
            }
            if (taxonomyFieldCity == null)
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
                        contentType.Update(false);
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
            }
            if (taxonomyFieldCity != null)
            {
                // Retrieve the field as a Taxonomy Field
                var taxColumn = context.CastTo<TaxonomyField>(taxonomyFieldCity);
                context.Load(taxColumn);
                context.ExecuteQuery();
                //initialize taxonomy field value
                var defaultValue = new TaxonomyFieldValue();
                defaultValue.WssId = -1;
                defaultValue.Label = "Ho Chi Minh";
                defaultValue.TermGuid = "d2ede839-c7df-4951-9eb6-6fc24f02cb44";
                //retrieve validated taxonomy field value
                var validatedValue = taxColumn.GetValidatedString(defaultValue);
                context.ExecuteQuery();
                //set default value for a taxonomy field
                taxColumn.DefaultValue = validatedValue.Value;
                taxColumn.Update();
                context.ExecuteQuery();
            }
            await context.ExecuteQueryAsync();
        }

        private static async Task CreateListAsync(ClientContext context, string listTitle)
        {
            ListCollection listCollection = context.Web.Lists;
            context.Load(listCollection, lists => lists.Include(list => list.Title).Where(list => list.Title == listTitle));
            await context.ExecuteQueryAsync();
            List list = null;
            if (listCollection.Count == 0)
            {
                Web web = context.Web;
                ListCreationInformation listCreationInformation = new ListCreationInformation();
                listCreationInformation.Title = listTitle;
                listCreationInformation.TemplateType = (int)ListTemplateType.GenericList;
                list = web.Lists.Add(listCreationInformation);
                await context.ExecuteQueryAsync();
            }
        }

        private static async Task AddListItemsAsync(ClientContext context, string listTitle)
        {
            ListCollection listCollection = context.Web.Lists;
            context.Load(listCollection, lists => lists.Include(list => list.Title).Where(list => list.Title == listTitle));
            await context.ExecuteQueryAsync();
            List list = null;
            if (listCollection.Count > 0)
            {
                list = listCollection.FirstOrDefault();
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
                context.ExecuteQuery();
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

        private static bool CheckIfListExists(ClientContext context, string listName)
        {
            ListCollection listCollection = context.Web.Lists;
            context.Load(listCollection, lists => lists.Include(list => list.Title).Where(list => list.Title == listName));
            context.ExecuteQuery();
            if (listCollection.Count == 0)
            {
                return false;
            }
            return true;
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

        private static async Task SimpleCamlQueryAsync(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle("Documents");

            var allItemsQuery = CamlQuery.CreateAllItemsQuery();
            var allFoldersQuery = CamlQuery.CreateAllFoldersQuery();

            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>
                                </Query>
                                <RowLimit>20</RowLimit>
                            </View>",
                FolderServerRelativeUrl = "/sites/test-site-duc-11111/Shared%20Documents/2"
                //example for site: https://omniapreprod.sharepoint.com/sites/test-site-duc-11111/
            });

            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
        }
    }
}
