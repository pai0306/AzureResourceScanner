import os
import pandas as pd
from azure.identity import ClientSecretCredential
from azure.mgmt.resource import ResourceManagementClient, SubscriptionClient
from azure.core.exceptions import ClientAuthenticationError

def scan_and_export_multi_tenant_resources():
    """
    主函式：使用單一服務主體憑證，
    遍歷設定好的多個租用戶，掃描所有可見資源，並匯出成單一 Excel 報告。
    """
    print("--- Azure 多租用戶資源權限掃描工具 (單一服務主體模式) ---")

    # --- 步驟 1: 從環境變數讀取服務主體憑證 ---
    # 這組憑證將被用來嘗試登入所有您指定的租用戶。
    try:
        client_id = os.environ["AZURE_CLIENT_ID"]
        client_secret = os.environ["AZURE_CLIENT_SECRET"]
        print(f"已成功讀取服務主體憑證 (Client ID: {client_id})。")
    except KeyError as e:
        print(f"錯誤：缺少環境變數 {e}。")
        print("請在執行前設定 AZURE_CLIENT_ID 和 AZURE_CLIENT_SECRET。")
        return

    # --- 步驟 2: 設定您要掃描的所有租用戶清單 ---
    # 您只需要提供一個好記的別名和對應的 Tenant ID。
    # 確保您使用的服務主體已被邀請並授權存取以下所有租用戶。
    TENANTS_TO_SCAN = [
        {
            "alias": "system A",
            "tenant_id": "",
        },
        {
            "alias": "system B",
            "tenant_id": "",
        },
        {
            "alias": "system C",
            "tenant_id": "",
        },
    ]

    output_excel_file = "Azure_Multi_Tenant_Resources_Report_Single_SP.xlsx"
    all_resources_data = []

    # --- 步驟 3: 遍歷每一個租用戶進行掃描 ---
    for tenant_info in TENANTS_TO_SCAN:
        alias = tenant_info["alias"]
        tenant_id = tenant_info["tenant_id"]

        print(f"\n==========================================================")
        print(f"正在處理租用戶: '{alias}' (Tenant ID: {tenant_id})")
        print(f"==========================================================")

        try:
            # --- 使用統一的憑證和當前迴圈的 tenant_id 進行驗證 ---
            print("正在進行驗證...")
            credential = ClientSecretCredential(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret
            )
            subscription_client = SubscriptionClient(credential)
            print("驗證成功！")

            print("正在獲取可存取的訂閱清單...")
            accessible_subscriptions = list(subscription_client.subscriptions.list())
            
            if not accessible_subscriptions:
                print("警告：在此租用戶中，此服務主體沒有權限存取任何訂閱。")
                continue

            print(f"找到 {len(accessible_subscriptions)} 個可存取的訂閱。")

            for sub in accessible_subscriptions:
                sub_id = sub.subscription_id
                sub_name = sub.display_name
                print(f"\n>>> 正在掃描訂閱: '{sub_name}' ({sub_id})")

                resource_client = ResourceManagementClient(credential, sub_id)
                resources_in_sub = list(resource_client.resources.list())

                if not resources_in_sub:
                    print(f" -> 在此訂閱中未找到任何資源。")
                    continue
                
                print(f" -> 找到 {len(resources_in_sub)} 個資源，正在處理...")
                for resource in resources_in_sub:
                    try:
                        resource_group_name = resource.id.split('/')[4]
                    except IndexError:
                        resource_group_name = "N/A (無法解析)"

                    tags_str = "; ".join([f"{key}={value}" for key, value in resource.tags.items()]) if resource.tags else ""

                    resource_info = {
                        "租用戶別名 (Tenant Alias)": alias,
                        "租用戶ID (Tenant ID)": tenant_id,
                        "訂閱名稱 (Subscription Name)": sub_name,
                        "訂閱ID (Subscription ID)": sub_id,
                        "資源群組 (Resource Group)": resource_group_name,
                        "資源名稱 (Resource Name)": resource.name,
                        "資源類型 (Resource Type)": resource.type,
                        "位置 (Location)": resource.location,
                        "標籤 (Tags)": tags_str
                    }
                    all_resources_data.append(resource_info)

        except ClientAuthenticationError:
            print(f"\n錯誤：租用戶 '{alias}' 的驗證失敗！")
            print("請確認此服務主體是否已被邀請至此租用戶，並且憑證有效。")
            continue
        except Exception as e:
            print(f"\n處理租用戶 '{alias}' 時發生未預期的錯誤: {e}")
            continue

    # --- 步驟 4: 將所有結果匯出至 Excel ---
    if not all_resources_data:
        print("\n掃描完成，但未在任何租用戶中找到可存取的資源。")
        return

    print(f"\n==========================================================")
    print(f"所有租用戶掃描完畢！總共收集到 {len(all_resources_data)} 筆資源資料。")
    print(f"正在將結果匯出至 Excel 檔案: {output_excel_file} ...")

    df = pd.DataFrame(all_resources_data)
    df = df[[
        "租用戶別名 (Tenant Alias)", "租用戶ID (Tenant ID)", "訂閱名稱 (Subscription Name)",
        "訂閱ID (Subscription ID)", "資源群組 (Resource Group)", "資源名稱 (Resource Name)",
        "資源類型 (Resource Type)", "位置 (Location)", "標籤 (Tags)"
    ]]
    
    df.to_excel(output_excel_file, index=False, engine='openpyxl')
    
    print(f"\n報告成功產出！請查看當前目錄下的檔案: {output_excel_file}")

if __name__ == "__main__":
    scan_and_export_multi_tenant_resources()
