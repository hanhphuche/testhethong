// Optimized version of the EPOModel processing code
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Data.Entity.Validation;
using Newtonsoft.Json;

public async Task<bool> ProcessEPOModelsOptimized(string result)
{
    var res = JsonConvert.DeserializeObject<List<EPOModel>>(result)
        ?.Where(x => !string.IsNullOrEmpty(x.AllPiCode))
        .ToList() ?? new List<EPOModel>();

    if (!res.Any())
        return true;

    DateTime dt_begin1 = DateTime.Now;

    try
    {
        // Batch collect all unique values to minimize database queries
        var uniqueValues = ExtractUniqueValues(res);
        
        // Execute all check queries in parallel
        var existingEntities = await GetExistingEntitiesAsync(uniqueValues);
        
        // Process missing entities in batches
        await CreateMissingEntitiesAsync(uniqueValues, existingEntities);
        
        // Refresh existing entities after creation
        existingEntities = await GetExistingEntitiesAsync(uniqueValues);
        
        // Process warehouse imports in batches
        await ProcessWarehouseImportsAsync(res, existingEntities);
        
        return true;
    }
    catch (Exception ex)
    {
        Log($"Lỗi xử lý EPO Models: {ex.Message}\n{ex.StackTrace}");
        throw;
    }
}

private class UniqueValues
{
    public HashSet<string> PoCodes { get; set; } = new HashSet<string>();
    public HashSet<string> ContractNos { get; set; } = new HashSet<string>();
    public HashSet<string> ProjectCodes { get; set; } = new HashSet<string>();
    public HashSet<string> ProducerNames { get; set; } = new HashSet<string>();
    public HashSet<string> SubDepartmentNames { get; set; } = new HashSet<string>();
    public HashSet<string> AllPiCodes { get; set; } = new HashSet<string>();
}

private class ExistingEntities
{
    public Dictionary<string, int> POIds { get; set; } = new Dictionary<string, int>();
    public Dictionary<string, int> SellContractIds { get; set; } = new Dictionary<string, int>();
    public Dictionary<string, int> ProjectIds { get; set; } = new Dictionary<string, int>();
    public Dictionary<string, int> ManufactorIds { get; set; } = new Dictionary<string, int>();
    public Dictionary<string, int> SubDepartmentIds { get; set; } = new Dictionary<string, int>();
    public Dictionary<string, int> BillNumIds { get; set; } = new Dictionary<string, int>();
}

private UniqueValues ExtractUniqueValues(List<EPOModel> items)
{
    var uniqueValues = new UniqueValues();
    
    foreach (var item in items)
    {
        if (!string.IsNullOrEmpty(item.PoCode))
            uniqueValues.PoCodes.Add(item.PoCode);
        if (!string.IsNullOrEmpty(item.ContractNo))
            uniqueValues.ContractNos.Add(item.ContractNo);
        if (!string.IsNullOrEmpty(item.ProjectCode))
            uniqueValues.ProjectCodes.Add(item.ProjectCode);
        if (!string.IsNullOrEmpty(item.ProducerName))
            uniqueValues.ProducerNames.Add(item.ProducerName);
        if (!string.IsNullOrEmpty(item.SubDepartmentName))
            uniqueValues.SubDepartmentNames.Add(item.SubDepartmentName);
        if (!string.IsNullOrEmpty(item.AllPiCode))
            uniqueValues.AllPiCodes.Add(item.AllPiCode);
    }
    
    return uniqueValues;
}

private async Task<ExistingEntities> GetExistingEntitiesAsync(UniqueValues uniqueValues)
{
    var existingEntities = new ExistingEntities();
    MSSQLHelper.DefaultConnectionString = CAMS;
    
    // Execute all queries in parallel
    var tasks = new List<Task>();
    
    if (uniqueValues.PoCodes.Any())
    {
        tasks.Add(Task.Run(() =>
        {
            var poInClause = string.Join("','", uniqueValues.PoCodes);
            var poSQL = $"SELECT PO_Name, ID FROM PO WHERE PO_Name IN ('{poInClause}')";
            
            var poResults = MSSQLHelper.ExecuteQueryListObject(poSQL, new object[0]);
            foreach (var result in poResults)
            {
                existingEntities.POIds[result[0].ToString()] = int.Parse(result[1].ToString());
            }
        }));
    }
    
    if (uniqueValues.ContractNos.Any())
    {
        tasks.Add(Task.Run(() =>
        {
            var contractInClause = string.Join("','", uniqueValues.ContractNos);
            var contractSQL = $"SELECT Name, ID FROM Sell_Contract WHERE Name IN ('{contractInClause}')";
            
            var contractResults = MSSQLHelper.ExecuteQueryListObject(contractSQL, new object[0]);
            foreach (var result in contractResults)
            {
                existingEntities.SellContractIds[result[0].ToString()] = int.Parse(result[1].ToString());
            }
        }));
    }
    
    if (uniqueValues.ProjectCodes.Any())
    {
        tasks.Add(Task.Run(() =>
        {
            var projectInClause = string.Join("','", uniqueValues.ProjectCodes);
            var projectSQL = $"SELECT ProjectCode, ID FROM ProjectCode WHERE ProjectCode IN ('{projectInClause}')";
            
            var projectResults = MSSQLHelper.ExecuteQueryListObject(projectSQL, new object[0]);
            foreach (var result in projectResults)
            {
                existingEntities.ProjectIds[result[0].ToString()] = int.Parse(result[1].ToString());
            }
        }));
    }
    
    if (uniqueValues.ProducerNames.Any())
    {
        tasks.Add(Task.Run(() =>
        {
            var manuInClause = string.Join("','", uniqueValues.ProducerNames);
            var manuSQL = $"SELECT Name, ID FROM Manufactor WHERE Name IN ('{manuInClause}')";
            
            var manuResults = MSSQLHelper.ExecuteQueryListObject(manuSQL, new object[0]);
            foreach (var result in manuResults)
            {
                existingEntities.ManufactorIds[result[0].ToString()] = int.Parse(result[1].ToString());
            }
        }));
    }
    
    if (uniqueValues.SubDepartmentNames.Any())
    {
        tasks.Add(Task.Run(() =>
        {
            var subDepInClause = string.Join("','", uniqueValues.SubDepartmentNames);
            var subDepSQL = $"SELECT Name, ID FROM warehouse_SubDepartment WHERE Name IN ('{subDepInClause}')";
            
            var subDepResults = MSSQLHelper.ExecuteQueryListObject(subDepSQL, new object[0]);
            foreach (var result in subDepResults)
            {
                existingEntities.SubDepartmentIds[result[0].ToString()] = int.Parse(result[1].ToString());
            }
        }));
    }
    
    if (uniqueValues.AllPiCodes.Any())
    {
        tasks.Add(Task.Run(() =>
        {
            var billInClause = string.Join("','", uniqueValues.AllPiCodes);
            var billSQL = $"SELECT BillNum, ID FROM BillNum WHERE BillNum IN ('{billInClause}')";
            
            var billResults = MSSQLHelper.ExecuteQueryListObject(billSQL, new object[0]);
            foreach (var result in billResults)
            {
                existingEntities.BillNumIds[result[0].ToString()] = int.Parse(result[1].ToString());
            }
        }));
    }
    
    await Task.WhenAll(tasks);
    return existingEntities;
}

private async Task CreateMissingEntitiesAsync(UniqueValues uniqueValues, ExistingEntities existingEntities)
{
    // Create missing POs
    var missingPOs = uniqueValues.PoCodes.Where(code => !existingEntities.POIds.ContainsKey(code)).ToList();
    if (missingPOs.Any())
    {
        try
        {
            var pos = missingPOs.Select(code => new PO
            {
                PO_Name = code,
                Created_Date = DateTime.Now,
                Created_By = "system"
            }).ToList();
            
            db.POes.AddRange(pos);
            db.SaveChanges();
        }
        catch (Exception ex)
        {
            Log("Lỗi Insert PO batch: " + ex.Message);
            throw;
        }
    }
    
    // Create missing Sell Contracts
    var missingSellContracts = uniqueValues.ContractNos.Where(code => !existingEntities.SellContractIds.ContainsKey(code)).ToList();
    if (missingSellContracts.Any())
    {
        try
        {
            var contracts = missingSellContracts.Select(code => new Sell_Contract
            {
                Name = code,
                Created_By = "system",
                Created_Date = DateTime.Now
            }).ToList();
            
            db.Sell_Contract.AddRange(contracts);
            db.SaveChanges();
        }
        catch (Exception ex)
        {
            Log("Lỗi Insert SellContract batch: " + ex.Message);
            throw;
        }
    }
    
    // Create missing Project Codes
    var missingProjectCodes = uniqueValues.ProjectCodes.Where(code => !existingEntities.ProjectIds.ContainsKey(code)).ToList();
    if (missingProjectCodes.Any())
    {
        try
        {
            var projects = missingProjectCodes.Select(code => new ProjectCode
            {
                ProjectCode1 = code,
                Created_By = "system",
                Created_Date = DateTime.Now
            }).ToList();
            
            db.ProjectCodes.AddRange(projects);
            db.SaveChanges();
        }
        catch (Exception ex)
        {
            Log("Lỗi Insert ProjectCode batch: " + ex.Message);
            throw;
        }
    }
    
    // Create missing Manufactors
    var missingManufactors = uniqueValues.ProducerNames.Where(code => !existingEntities.ManufactorIds.ContainsKey(code)).ToList();
    if (missingManufactors.Any())
    {
        try
        {
            var manufactors = missingManufactors.Select(name => new Manufactor
            {
                Name = name,
                Created_By = "system",
                Created_Date = DateTime.Now
            }).ToList();
            
            db.Manufactors.AddRange(manufactors);
            db.SaveChanges();
        }
        catch (Exception ex)
        {
            Log("Lỗi Insert Manufactor batch: " + ex.Message);
            throw;
        }
    }
    
    // Create missing SubDepartments
    var missingSubDeps = uniqueValues.SubDepartmentNames.Where(code => !existingEntities.SubDepartmentIds.ContainsKey(code)).ToList();
    if (missingSubDeps.Any())
    {
        try
        {
            var subDeps = missingSubDeps.Select(name => new warehouse_SubDepartment
            {
                Name = name,
                created_by = "system",
                created_at = DateTime.Now,
                description = "Created by system"
            }).ToList();
            
            db.warehouse_SubDepartment.AddRange(subDeps);
            db.SaveChanges();
        }
        catch (Exception ex)
        {
            Log("Lỗi Insert SubDepartment batch: " + ex.Message);
            throw;
        }
    }
    
    // Create missing BillNums
    var missingBillNums = uniqueValues.AllPiCodes.Where(code => !existingEntities.BillNumIds.ContainsKey(code)).ToList();
    if (missingBillNums.Any())
    {
        try
        {
            var billNums = missingBillNums.Select(code => new BillNum
            {
                BillNum1 = code,
                Created_By = "system",
                Created_Date = DateTime.Now
            }).ToList();
            
            db.BillNums.AddRange(billNums);
            db.SaveChanges();
        }
        catch (Exception ex)
        {
            Log("Lỗi Insert BillNum batch: " + ex.Message);
            throw;
        }
    }
}

private async Task ProcessWarehouseImportsAsync(List<EPOModel> items, ExistingEntities existingEntities)
{
    const int batchSize = 1000;
    
    // Group items by PO and BillNum for efficient processing
    var groupedItems = items.GroupBy(x => new { x.PoCode, x.AllPiCode, x.PoItemId }).ToList();
    
    foreach (var batch in groupedItems.Batch(batchSize))
    {
        var warehouseImportsToAdd = new List<warehouse_import>();
        var warehouseImportsToRemove = new List<warehouse_import>();
        
        foreach (var group in batch)
        {
            var item = group.First(); // Take first item as template
            var quantity = group.Sum(x => x.Quantity);
            
            // Get existing products for this group
            var poId = existingEntities.POIds[item.PoCode];
            var cleanItemName = item.ItemName.Replace("\r", "").Replace("\n", "").Trim();
            
            var existedProducts = db.warehouse_import
                .Where(x => x.pro_name == cleanItemName && 
                           x.PO_NumID == poId && 
                           x.BillNum == item.AllPiCode && 
                           x.PoItemId == item.PoItemId)
                .ToList();
            
            if (existedProducts.Any())
            {
                // Update existing products
                foreach (var existedProduct in existedProducts)
                {
                    UpdateWarehouseImportFromItem(existedProduct, item, existingEntities);
                }
                
                // Handle quantity differences
                var quantityDiff = quantity - existedProducts.Count;
                if (quantityDiff > 0)
                {
                    // Add more items
                    var template = existedProducts.First();
                    for (int i = 0; i < quantityDiff; i++)
                    {
                        warehouseImportsToAdd.Add(CreateWarehouseImportFromTemplate(template));
                    }
                }
                else if (quantityDiff < 0)
                {
                    // Remove excess items (only those without serial numbers)
                    var toRemove = existedProducts
                        .Where(x => string.IsNullOrEmpty(x.Serial_num))
                        .Take(Math.Abs(quantityDiff))
                        .ToList();
                    warehouseImportsToRemove.AddRange(toRemove);
                }
            }
            else
            {
                // Create new warehouse imports
                for (int i = 0; i < quantity; i++)
                {
                    warehouseImportsToAdd.Add(CreateWarehouseImportFromItem(item, existingEntities));
                }
            }
        }
        
        // Execute batch operations
        try
        {
            if (warehouseImportsToAdd.Any())
            {
                db.warehouse_import.AddRange(warehouseImportsToAdd);
            }
            
            if (warehouseImportsToRemove.Any())
            {
                db.warehouse_import.RemoveRange(warehouseImportsToRemove);
            }
            
            db.SaveChanges();
        }
        catch (DbEntityValidationException dbEx)
        {
            foreach (var validationErrors in dbEx.EntityValidationErrors)
            {
                foreach (var validationError in validationErrors.ValidationErrors)
                {
                    Log($"Property: {validationError.PropertyName} Error: {validationError.ErrorMessage}");
                }
            }
            throw;
        }
        catch (Exception ex)
        {
            Log("Lỗi batch update warehouse_import: " + ex.Message);
            throw;
        }
    }
}

private void UpdateWarehouseImportFromItem(warehouse_import warehouseImport, EPOModel item, ExistingEntities existingEntities)
{
    warehouseImport.AM = item.AmAccount;
    warehouseImport.FRU = item.PartNo;
    warehouseImport.ProjectID = existingEntities.ProjectIds[item.ProjectCode];
    warehouseImport.manu_Bought_time = item.Guarantee.ToString();
    warehouseImport.ManufactorID = existingEntities.ManufactorIds[item.ProducerName];
    warehouseImport.BP_Account = item.PoMan;
    warehouseImport.Sell_ContractID = existingEntities.SellContractIds[item.ContractNo];
    warehouseImport.PoItemId = item.PoItemId;
    warehouseImport.BillNum = item.AllPiCode;
    warehouseImport.SubDepId = existingEntities.SubDepartmentIds[item.SubDepartmentName];
}

private warehouse_import CreateWarehouseImportFromItem(EPOModel item, ExistingEntities existingEntities)
{
    return new warehouse_import
    {
        ManufactorID = existingEntities.ManufactorIds[item.ProducerName],
        ProjectID = existingEntities.ProjectIds[item.ProjectCode],
        PO_NumID = existingEntities.POIds[item.PoCode],
        pro_name = item.ItemName.Replace("\r", "").Replace("\n", "").Trim(),
        FRU = item.PartNo,
        BP_Account = item.PoMan,
        Sell_ContractID = existingEntities.SellContractIds[item.ContractNo],
        manu_Bought_time = item.Guarantee.ToString(),
        AM = item.AmAccount,
        BillNum = item.AllPiCode,
        synced_through_epo = true,
        Exported = 0,
        Created_Date = DateTime.Now,
        PoItemId = item.PoItemId,
        SubDepId = existingEntities.SubDepartmentIds[item.SubDepartmentName]
    };
}

private warehouse_import CreateWarehouseImportFromTemplate(warehouse_import template)
{
    return new warehouse_import
    {
        ManufactorID = template.ManufactorID,
        ProjectID = template.ProjectID,
        PO_NumID = template.PO_NumID,
        pro_name = template.pro_name,
        FRU = template.FRU,
        BP_Account = template.BP_Account,
        Sell_ContractID = template.Sell_ContractID,
        manu_Bought_time = template.manu_Bought_time,
        AM = template.AM,
        BillNum = template.BillNum,
        synced_through_epo = true,
        Exported = 0,
        Created_Date = DateTime.Now,
        PoItemId = template.PoItemId,
        SubDepId = template.SubDepId
    };
}

// Extension method for batching
public static class EnumerableExtensions
{
    public static IEnumerable<IEnumerable<T>> Batch<T>(this IEnumerable<T> source, int batchSize)
    {
        using (var enumerator = source.GetEnumerator())
        {
            while (enumerator.MoveNext())
            {
                yield return YieldBatchElements(enumerator, batchSize - 1);
            }
        }
    }

    private static IEnumerable<T> YieldBatchElements<T>(IEnumerator<T> source, int batchSize)
    {
        yield return source.Current;
        for (int i = 0; i < batchSize && source.MoveNext(); i++)
        {
            yield return source.Current;
        }
    }
}