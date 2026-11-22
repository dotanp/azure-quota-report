# Azure vCPU quota report

1. Install Python 3.9+ and [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli).
2. Copy [quota_input_template.xlsx](./quota_input_template.xlsx) or edit [quota_input_example.xlsx](./quota_input_example.xlsx). Replace the placeholder subscription IDs in the example.
3. Run:

   ```sh
   python3 -m pip install -r requirements.txt
   az login
   az extension add --name quota
   python3 azure_quota_report.py quota_input_example.xlsx --output quota_report.xlsx
   ```

Use a virtual environment if your Python installation does not allow system-wide packages. Your Azure identity must have read access to each subscription in the workbook. The script uses your existing Azure CLI sign-in, checks each subscription explicitly, and does not submit quota requests or modify Azure resources.

Keep workbooks containing real subscription IDs and quota data out of version control. The repository tracks only the blank templates and the example with placeholder IDs. Save your own input under a different name so Git ignores it.

Enter one request per row on the **Requests** tab. Choose a **Region** code such as `eastus` from the dropdown. The **Azure regions** tab maps codes to full names using Microsoft's [public Azure regions list](https://learn.microsoft.com/en-us/azure/reliability/regions-list). The dropdown blocks unknown codes, but a listed region might not be available to your subscription.

Choose a **Service Name** from the dropdown. The dropdown uses a snapshot of Microsoft's [public Azure product catalog](https://azure.microsoft.com/en-us/products/) and blocks unknown names typed into the cell. The **Azure services** tab shows which entries are supported: this report checks **Virtual Machines** and **Virtual Machine Scale Sets** only. Other selections appear on the **Issues** tab, and existing inputs using `Azure VM` or `VMSS` still work.

**Required cores** is the total *additional* vCPUs for the row, not a VM count. For example, four 4-vCPU VMs require 16 cores.

## SKU format and lookup

In the **SKU** column, enter the exact VM size name, including the `Standard_` prefix and underscores. For example, enter `Standard_D4ds_v6` or `Standard_D8ds_v6`, depending on the size you plan to deploy. Do not enter the portal quota label `Standard Ddsv6 Family vCPUs` or the family name `Ddsv6`. The script looks up the quota family from the VM size.

After `az login`, replace `YOUR_SUBSCRIPTION_ID` and `eastus` with the values from your Excel row. Copy a value from the **SKU** column:

```sh
az vm list-skus --subscription YOUR_SUBSCRIPTION_ID --location eastus \
  --resource-type virtualMachines --query '[].{SKU:name,Family:family}' -o table
```

Add `--all` to include SKUs that aren't available to the subscription. To list VM SKUs across regions, omit `--location`:

```sh
az vm list-skus --subscription YOUR_SUBSCRIPTION_ID --resource-type virtualMachines \
  --all --query '[].{SKU:name,Family:family}' -o table
```

These commands list VM SKUs, not SKU names for other Azure services. A listed SKU does not guarantee regional capacity.

## Reading the output

- **Quota increase needed (y/n):** `y` if either the VM-family or total regional vCPU limit is too low, `n` if both are sufficient, and blank if the script couldn't determine the answer.
- **current quota:** The VM-family limit in vCPUs, not the number of vCPUs still available.
- **increase to:** The minimum new limit for each quota that needs an increase: current usage plus the total planned vCPUs sharing that quota. If several rows share a quota, the same target is displayed on each row; do not add the targets together.

Open the output in Excel to calculate the formulas. **Quota details** shows the limits, usage, planned cores, and targets for each quota. The script first reads limits and usage with `az quota` and falls back to `az vm list-usage` when that quota is missing or incomplete. **Issues** lists rows that couldn't be fully verified; the script still writes the workbook but exits with status 1. Correct those rows and rerun before relying on totals.

This checks standard (non-Spot) Azure VM and scale-set vCPUs only, not service-specific quotas, VM counts, SKU availability, or regional hardware capacity. Edit the input workbook and rerun the script when the deployment plan changes or to refresh the Azure snapshot. The selected output path is overwritten if it already exists, but it cannot be the input path.

See [quota_output_template.xlsx](./quota_output_template.xlsx) for the report layout. Use `python3 azure_quota_report.py --make-templates` to regenerate (and overwrite) the three template/example files in the current directory. This downloads the latest public service and region lists and does not require Azure sign-in.
