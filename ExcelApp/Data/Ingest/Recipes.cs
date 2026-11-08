using System.Text.Json;
using System.Text.RegularExpressions;
namespace ExcelApp.Data;
public sealed class Recipes
{
    readonly ConfigRoot root; readonly object gate=new();
    public Recipes(ConfigRoot root) {this.root=root;}
    public IReadOnlyList<IngestRecipe> List() {lock(gate) return root.List("recipes","*.json").Select(p=>JsonSerializer.Deserialize<IngestRecipe>(root.ReadAllText(p),Json.Options)!).Where(x=>x!=null).ToArray();}
    public void Save(IngestRecipe recipe) {
        recipe.Id=Regex.Replace(recipe.Id.ToLowerInvariant(),"[^a-z0-9]+","-").Trim('-');
        if(recipe.Id.Length==0 || recipe.Folders.Count==0 || recipe.Patterns.Count==0 || recipe.Folders.Any(string.IsNullOrWhiteSpace) || recipe.Patterns.Any(string.IsNullOrWhiteSpace)) throw new UserError("Choose a name, folder and file pattern.");
        if(!recipe.HeaderSignature.Contains(recipe.DateColumn,StringComparer.OrdinalIgnoreCase) || recipe.KeyColumns.Count==0 || recipe.KeyColumns.Any(k=>!recipe.HeaderSignature.Contains(k,StringComparer.OrdinalIgnoreCase))) throw new UserError("The date and key columns must be in the signature.");
        if(recipe.PartialThreshold<0 || recipe.PartialThreshold>0.9 || double.IsNaN(recipe.PartialThreshold)) throw new UserError("The partial-file threshold must be between 0 and 0.9.");
        lock(gate) root.WriteAllText($"recipes/{recipe.Id}.json",JsonSerializer.Serialize(recipe,Json.Options));
    }
    public void Delete(string id) {if(!Regex.IsMatch(id,"^[a-z0-9-]+$")) throw new UserError("Invalid recipe.");lock(gate) root.Delete($"recipes/{id}.json");}
}
