 (cd "$(git rev-parse --show-toplevel)" && git apply --3way <<'EOF' 
diff --git a/Shopify Bulk Tool.py	 b/Shopify Bulk Tool.py	
index a227c98f97a01d759c62ab43087c8435f4dc01b9..9313edb90ff82dc43ff39a98099d8fed0b322eb0 100644
--- a/Shopify Bulk Tool.py	
+++ b/Shopify Bulk Tool.py	
@@ -130,51 +130,52 @@ def read_credentials(file_path):
     credentials = {}
     with open(file_path, 'r') as file:
         for line in file:
             key, value = line.strip().split('=')
             credentials[key] = value
     return credentials
 
 
 
 def run_downloader_logic():
     # Get the directory where the executable or script is located
     if getattr(sys, 'frozen', False):  # If running as an EXE
         script_dir = os.path.dirname(sys.executable)
     else:
         script_dir = os.path.dirname(os.path.abspath(__file__))
 
     # Build the full path to 'credentials.txt'
     credentials_path = os.path.join(script_dir, 'credentials.txt')
 
     # Load credentials from 'credentials.txt'
     credentials = read_credentials(credentials_path)
     SHOP_NAME = credentials['store_name']
     ACCESS_TOKEN = credentials['access_token']
 
     # Shopify API URL
-    BASE_URL = f"https://{SHOP_NAME}.myshopify.com/admin/api/2023-07"
+    BASE_URL = f"https://{SHOP_NAME}.myshopify.com/admin/api/2023-07"
+    GRAPHQL_URL = f"https://{SHOP_NAME}.myshopify.com/admin/api/2023-07/graphql.json"
 
     # Headers for API authentication
     headers = {
         "Content-Type": "application/json",
         "X-Shopify-Access-Token": ACCESS_TOKEN
     }
 
     # The rest of your downloader logic goes here...
 
     # Fetch all locations
     def get_locations():
         url = f"{BASE_URL}/locations.json"
         response = requests.get(url, headers=headers)
         
         if response.status_code == 200:
             return response.json().get('locations', [])
         return []
 
     # Fetch inventory levels for a specific variant at a location
     def get_inventory_level(inventory_item_id, location_id):
         url = f"{BASE_URL}/inventory_levels.json?inventory_item_ids={inventory_item_id}&location_ids={location_id}"
         response = requests.get(url, headers=headers)
         
         if response.status_code == 200:
             inventory_levels = response.json().get('inventory_levels', [])
diff --git a/Shopify Bulk Tool.py	 b/Shopify Bulk Tool.py	
index a227c98f97a01d759c62ab43087c8435f4dc01b9..9313edb90ff82dc43ff39a98099d8fed0b322eb0 100644
--- a/Shopify Bulk Tool.py	
+++ b/Shopify Bulk Tool.py	
@@ -205,65 +206,152 @@ def run_downloader_logic():
             products.extend(batch)
 
             # Properly handle pagination if more products are available
             link_header = response.headers.get('Link')
             next_url = None
             if link_header:
                 links = link_header.split(',')
                 for link in links:
                     if 'rel="next"' in link:
                         next_url = link.split(';')[0].strip('<> ').replace('&amp;', '&')
                         break
 
             if next_url:
                 url = next_url
                 page += 1
             else:
                 url = None
 
         print(f"Finished fetching products. Total products fetched: {len(products)}")
         return products
 
 
 
 
     # Fetch metafields for a specific product
-    def get_metafields(owner_id, owner_resource="product"):
-        print(f"Fetching metafields for {owner_resource} ID {owner_id}...")
-        time.sleep(1)  # Wait for 1 second after every request
+    def get_metafields(owner_id, owner_resource="product"):
+        print(f"Fetching metafields for {owner_resource} ID {owner_id}...")
+        time.sleep(1)  # Wait for 1 second after every request
 
 
         url = f"{BASE_URL}/metafields.json?metafield[owner_id]={owner_id}&metafield[owner_resource]={owner_resource}"
         response = requests.get(url, headers=headers)
 
         if response.status_code == 200:
             return response.json().get('metafields', [])
         else:
             print(f"Error fetching metafields for {owner_resource} ID {owner_id}: {response.status_code}")
         return []
 
-    def get_image_url_from_gid(gid):
+    file_info_cache = {}
+
+    def fetch_file_info_for_gid(gid):
+        if not gid or not isinstance(gid, str):
+            return {"alt": "", "url": "", "filename": ""}
+
+        if gid in file_info_cache:
+            return file_info_cache[gid]
+
+        query = {
+            "query": f"""
+            {{
+                node(id: "{gid}") {{
+                    id
+                    ... on MediaImage {{
+                        alt
+                        image {{
+                            url
+                        }}
+                    }}
+                    ... on GenericFile {{
+                        alt
+                        url
+                    }}
+                }}
+            }}
+            """
+        }
+
+        alt_text = ""
+        file_url = ""
+        filename = ""
+
+        try:
+            response = requests.post(GRAPHQL_URL, json=query, headers=headers)
+            if response.status_code == 200:
+                data = response.json().get("data", {})
+                node = data.get("node") if data else None
+                if node:
+                    alt_text = node.get("alt", "") or ""
+                    if node.get("url"):
+                        file_url = node.get("url") or ""
+                    else:
+                        image_data = node.get("image") or {}
+                        file_url = image_data.get("url", "") if isinstance(image_data, dict) else ""
+            else:
+                print(f"Failed to fetch file info for gid {gid}: {response.status_code}, {response.text}")
+        except Exception as exc:
+            print(f"Error fetching file info for gid {gid}: {exc}")
+
+        if file_url:
+            try:
+                filename = os.path.basename(urllib.parse.urlparse(file_url).path)
+            except Exception:
+                filename = os.path.basename(file_url)
+
+        info = {"alt": alt_text, "url": file_url, "filename": filename}
+        file_info_cache[gid] = info
+        return info
+
+    def extract_file_reference_gids(raw_value):
+        if raw_value is None:
+            return []
+        if isinstance(raw_value, list):
+            return [str(item) for item in raw_value if item]
+        if isinstance(raw_value, str):
+            value = raw_value.strip()
+            if not value:
+                return []
+            try:
+                parsed = json.loads(value)
+                if isinstance(parsed, list):
+                    return [str(item) for item in parsed if item]
+                if isinstance(parsed, str):
+                    return [parsed]
+            except json.JSONDecodeError:
+                return [item.strip() for item in value.split(',') if item.strip()]
+            return [value]
+        return []
+
+    def format_alt_texts_for_storage(alt_texts, is_list=False):
+        if not alt_texts:
+            return json.dumps([]) if is_list else ""
+        if is_list:
+            return json.dumps([text or "" for text in alt_texts], ensure_ascii=False)
+        return alt_texts[0] if alt_texts else ""
+
+    def get_image_url_from_gid(gid):
         query = {
             "query": f"""
             {{
                 media(id: "{gid}") {{
                     ... on MediaImage {{
                         image {{
                             originalSrc
                         }}
                     }}
                 }}
             }}
             """
         }
         url = f"https://{SHOP_NAME}.myshopify.com/admin/api/2023-07/graphql.json"
         response = requests.post(url, json=query, headers=headers)
         if response.status_code == 200:
             data = response.json()
             return data['data']['media']['image']['originalSrc']
         else:
             print(f"Failed to fetch image URL for gid {gid}: {response.status_code}")
             return None
   
     def fetch_all_metafields(products):
         print("Fetching metafields for all products concurrently...")
         product_id_to_metafields = {}
diff --git a/Shopify Bulk Tool.py	 b/Shopify Bulk Tool.py	
index a227c98f97a01d759c62ab43087c8435f4dc01b9..9313edb90ff82dc43ff39a98099d8fed0b322eb0 100644
--- a/Shopify Bulk Tool.py	
+++ b/Shopify Bulk Tool.py	
@@ -420,70 +508,89 @@ def run_downloader_logic():
                 "Variant Compare At Price": first_variant.get('compare_at_price', ''),
                 "Variant Inventory Qty": first_variant.get('inventory_quantity', 0),
                 "Variant Weight": first_variant.get('weight', ''),
                 "Variant Weight Unit": first_variant.get('weight_unit', ''),
                 "Variant Barcode": first_variant.get('barcode', ''),
                 "Continue Selling When Sold Out": first_variant.get('inventory_policy', ''),
                 "Variant Image": variant_image_url,
                 "Variant Image Alt": variant_image_alt,
                 "Status": product['status'],  # Include status directly in the product data
 
             }
 
             # Add image URLs and alt texts dynamically after the variant data
             for i in range(max_image_columns):
                 if i < len(images):
                     image_url = images[i]['src']
                     image_alt = images[i].get('alt', '')
                 else:
                     image_url = None
                     image_alt = None
                 product_data[f"Image {i + 1}"] = image_url
                 product_data[f"Image {i + 1} Alt"] = image_alt
 
             
              # Add product-level metafields after the images
-            metafields = product_id_to_metafields.get(product['id'], [])
-            for metafield in metafields:
-                key = metafield['key']
-                value = metafield['value']
-                namespace = metafield['namespace']
-                field_type = metafield.get('type', 'unknown')
-                column_name = f"Metafield: {namespace}.{key} [{field_type}]"
-                all_metafield_keys.add(column_name)
-                product_data[column_name] = value
-
-            # Add variant-level metafields for the first variant
-            first_variant_metafields = variant_id_to_metafields.get(first_variant.get('id'), [])
-            for metafield in first_variant_metafields:
-                key = metafield['key']
-                value = metafield['value']
-                namespace = metafield['namespace']
-                field_type = metafield.get('type', 'unknown')
-                column_name = f"Variant Metafield: {namespace}.{key} [{field_type}]"
-                all_metafield_keys.add(column_name)
-                product_data[column_name] = value
+            metafields = product_id_to_metafields.get(product['id'], [])
+            for metafield in metafields:
+                key = metafield['key']
+                value = metafield['value']
+                namespace = metafield['namespace']
+                field_type = metafield.get('type', 'unknown')
+                column_name = f"Metafield: {namespace}.{key} [{field_type}]"
+                all_metafield_keys.add(column_name)
+                product_data[column_name] = value
+
+                if field_type in ('file_reference', 'list.file_reference'):
+                    gids = extract_file_reference_gids(value)
+                    alt_texts = [fetch_file_info_for_gid(gid).get('alt', '') for gid in gids]
+                    alt_column_name = f"Metafield Alt: {namespace}.{key}"
+                    all_metafield_keys.add(alt_column_name)
+                    product_data[alt_column_name] = format_alt_texts_for_storage(
+                        alt_texts,
+                        is_list=(field_type == 'list.file_reference')
+                    )
+
+            # Add variant-level metafields for the first variant
+            first_variant_metafields = variant_id_to_metafields.get(first_variant.get('id'), [])
+            for metafield in first_variant_metafields:
+                key = metafield['key']
+                value = metafield['value']
+                namespace = metafield['namespace']
+                field_type = metafield.get('type', 'unknown')
+                column_name = f"Variant Metafield: {namespace}.{key} [{field_type}]"
+                all_metafield_keys.add(column_name)
+                product_data[column_name] = value
+                if field_type in ('file_reference', 'list.file_reference'):
+                    gids = extract_file_reference_gids(value)
+                    alt_texts = [fetch_file_info_for_gid(gid).get('alt', '') for gid in gids]
+                    alt_column_name = f"Variant Metafield Alt: {namespace}.{key}"
+                    all_metafield_keys.add(alt_column_name)
+                    product_data[alt_column_name] = format_alt_texts_for_storage(
+                        alt_texts,
+                        is_list=(field_type == 'list.file_reference')
+                    )
 
             # Get inventory levels for each location for the first variant
             for location in locations:
                 location_id = location['id']
                 location_name = location['name']
                 key = (first_variant.get('inventory_item_id'), location_id)
                 inventory_level = inventory_mapping.get(key, 0)
                 product_data[f"Inventory Available: {location_name}"] = inventory_level
 
             # Append the first row for the product
             data.append(product_data)
 
             # Additional rows for the rest of the variants (without product-level details)
             variant_count = len(product.get('variants', []))
             for v_idx, variant in enumerate(product.get('variants', [])[1:], 2):
                 print(f"  Processing variant {v_idx}/{variant_count}: {variant.get('sku', '')}")
                 variant_id = variant.get('id', '')
                 variant_image_url = ''
                 variant_image_alt = ''
                 image_id = variant.get('image_id')
                 if image_id and image_id in image_id_to_image:
                     variant_image_url = image_id_to_image[image_id]['src']
                     variant_image_alt = image_id_to_image[image_id]['alt']
 
                 variant_data = {
diff --git a/Shopify Bulk Tool.py	 b/Shopify Bulk Tool.py	
index a227c98f97a01d759c62ab43087c8435f4dc01b9..9313edb90ff82dc43ff39a98099d8fed0b322eb0 100644
--- a/Shopify Bulk Tool.py	
+++ b/Shopify Bulk Tool.py	
@@ -493,59 +600,68 @@ def run_downloader_logic():
                     "Option1 Value": variant.get('option1', '') if len(product.get('options', [])) > 0 else "",
                     "Option2 Name": product.get('options', [{}])[1].get('name', '') if len(product.get('options', [])) > 1 else "",
                     "Option2 Value": variant.get('option2', '') if len(product.get('options', [])) > 1 else "",
                     "Option3 Name": product.get('options', [{}])[2].get('name', '') if len(product.get('options', [])) > 2 else "",
                     "Option3 Value": variant.get('option3', '') if len(product.get('options', [])) > 2 else "",
                     "Variant SKU": variant.get('sku', ''),
                     "Variant Price": variant.get('price', ''),
                     "Variant Compare At Price": variant.get('compare_at_price', ''),
                     "Variant Inventory Qty": variant.get('inventory_quantity', 0),
                     "Variant Weight": variant.get('weight', ''),
                     "Variant Weight Unit": variant.get('weight_unit', ''),
                     "Variant Barcode": variant.get('barcode', ''),
                     "Continue Selling When Sold Out": variant.get('inventory_policy', ''),
                     "Variant Image": variant_image_url,
                     "Variant Image Alt": variant_image_alt
                 }
 
                 # Get inventory levels for each location
                 for location in locations:
                     location_id = location['id']
                     location_name = location['name']
                     key = (variant.get('inventory_item_id'), location_id)
                     inventory_level = inventory_mapping.get(key, 0)
                     variant_data[f"Inventory Available: {location_name}"] = inventory_level
 
-                variant_metafields = variant_id_to_metafields.get(variant.get('id'), [])
-                for metafield in variant_metafields:
-                    key = metafield['key']
-                    value = metafield['value']
-                    namespace = metafield['namespace']
-                    field_type = metafield.get('type', 'unknown')
-                    column_name = f"Variant Metafield: {namespace}.{key} [{field_type}]"
-                    all_metafield_keys.add(column_name)
-                    variant_data[column_name] = value
+                variant_metafields = variant_id_to_metafields.get(variant.get('id'), [])
+                for metafield in variant_metafields:
+                    key = metafield['key']
+                    value = metafield['value']
+                    namespace = metafield['namespace']
+                    field_type = metafield.get('type', 'unknown')
+                    column_name = f"Variant Metafield: {namespace}.{key} [{field_type}]"
+                    all_metafield_keys.add(column_name)
+                    variant_data[column_name] = value
+                    if field_type in ('file_reference', 'list.file_reference'):
+                        gids = extract_file_reference_gids(value)
+                        alt_texts = [fetch_file_info_for_gid(gid).get('alt', '') for gid in gids]
+                        alt_column_name = f"Variant Metafield Alt: {namespace}.{key}"
+                        all_metafield_keys.add(alt_column_name)
+                        variant_data[alt_column_name] = format_alt_texts_for_storage(
+                            alt_texts,
+                            is_list=(field_type == 'list.file_reference')
+                        )
 
                 # Append the variant row
                 data.append(variant_data)
 
             # Append a blank row for separation (optional)
             data.append({})  # Adding an empty row for grouping in Excel
 
         # Create DataFrame and ensure all metafield columns are present
         df = pd.DataFrame(data)
         for key in all_metafield_keys:
             if key not in df.columns:
                 df[key] = None
 
         # Get the current date and time to append to the filename
         current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
         file_path = f"shopify_products_bulk_{current_time}.xlsx"
 
         # Save to Excel
         df.to_excel(file_path, index=False)
 
         # Open the saved Excel file using openpyxl to apply freezing panes and bold styling
         wb = load_workbook(file_path)
         ws = wb.active
 
         # Apply bold to product rows (rows where "ID" is present)
diff --git a/Shopify Bulk Tool.py	 b/Shopify Bulk Tool.py	
index a227c98f97a01d759c62ab43087c8435f4dc01b9..9313edb90ff82dc43ff39a98099d8fed0b322eb0 100644
--- a/Shopify Bulk Tool.py	
+++ b/Shopify Bulk Tool.py	
@@ -1054,358 +1170,393 @@ def run_uploader_logic():
         data = response.json()
         if "data" in data and "stagedUploadsCreate" in data["data"]:
             return data["data"]["stagedUploadsCreate"]["stagedTargets"][0]
         else:
             print("Error in staged upload creation:", data)
             return None
 
     # Function to upload the file to the staging URL
     def upload_file_to_staging(staging_target, file_path):
         url = staging_target["url"]
         files = {"file": open(file_path, "rb")}
         form_data = {param["name"]: param["value"] for param in staging_target["parameters"]}
         response = requests.post(url, data=form_data, files=files)
 
         # Parsing the XML response to get the file location
         if response.status_code == 201:
             print(f"File {file_path} successfully uploaded to staging URL.")
             xml_response = ET.fromstring(response.text)
             location_url = xml_response.find('Location').text  # Extracting Location URL from XML
             return location_url
         else:
             print(f"Failed to upload file. Status: {response.status_code}, {response.text}")
             return None
 
     # Function to commit the file with fileCreate
-    def commit_file_to_shopify(file_name, original_source):
-        query = """
-        mutation fileCreate($files: [FileCreateInput!]!) {
-        fileCreate(files: $files) {
-            files {
-            id  # Fetch the gid after file creation
+    def commit_file_to_shopify(file_name, original_source, alt_text=None):
+        query = """
+        mutation fileCreate($files: [FileCreateInput!]!) {
+        fileCreate(files: $files) {
+            files {
+            id  # Fetch the gid after file creation
             alt
             createdAt
             ... on GenericFile {
                 url
             }
             ... on MediaImage {
                 image {
                 url
                 }
             }
             }
             userErrors {
             code
             field
             message
             }
         }
-        }
-        """
-        variables = {
-            "files": [
-                {
-                    "alt": file_name,
-                    "originalSource": original_source
-                }
-            ]
-        }
-        response = requests.post(GRAPHQL_URL, json={"query": query, "variables": variables}, headers=graphql_headers)
+        }
+        """
+        variables = {
+            "files": [
+                {
+                    "alt": alt_text if alt_text is not None else file_name,
+                    "originalSource": original_source
+                }
+            ]
+        }
+        response = requests.post(GRAPHQL_URL, json={"query": query, "variables": variables}, headers=graphql_headers)
         data = response.json()
         if "data" in data and "fileCreate" in data["data"]:
             file_info = data["data"]["fileCreate"]["files"]
             gid = file_info[0]["id"]
             print(f"File {file_name} successfully committed to Shopify.")
             print(f"GID: {gid}")  # Print the gid here
             return file_info
         else:
             print("Error in file commit:", data)
             return None
 
     # Function to get image URL for a given GID
     def get_image_url_for_gid(gid):
         query = f"""
         {{
         node(id: "{gid}") {{
             ... on GenericFile {{
             url
             }}
             ... on MediaImage {{
             image {{
                 url
             }}
             }}
         }}
         }}
         """
         response = requests.post(GRAPHQL_URL, json={"query": query}, headers=graphql_headers)
         if response.status_code == 200:
             data = response.json()
             node = data.get('data', {}).get('node', {})
             url = None
             if 'url' in node:
                 url = node['url']
             elif 'image' in node and node['image']:
                 url = node['image'].get('url')
             return url
         else:
             print(f"Failed to get image URL for GID {gid}")
             return None
 
    
         
     # Function to upload an image to Shopify
-    def upload_image_to_shopify(file_path):
-        filename = os.path.basename(file_path)
-
-        # Normalize the filename before accessing it
-        normalized_filename = normalize_filename(filename)
-        folder_path = os.path.dirname(file_path)
+    def upload_image_to_shopify(file_path, alt_text=None):
+        filename = os.path.basename(file_path)
+
+        # Normalize the filename before accessing it
+        normalized_filename = normalize_filename(filename)
+        folder_path = os.path.dirname(file_path)
 
         file_path = os.path.join(os.path.dirname(file_path), normalized_filename)
 
         
         # Check if file exists after normalization
         if not os.path.exists(file_path):
             print(f"Image file {normalized_filename} not found in local folder.")
             return None, None
 
 
           # Ensure the filename is URL-encoded
         encoded_filename = encode_filename(filename)    
     
     
         # Extract the file name and MIME type
         mime_type = "image/jpeg" if filename.lower().endswith(".jpg") else "image/png"
 
         # Step 1: Create the staged upload
         staging_target = staged_upload_create(filename, mime_type)
         if not staging_target:
             return None, None
 
         # Step 2: Upload the file to the staging URL
         location_url = upload_file_to_staging(staging_target, file_path)
         if not location_url:
             return None, None
 
-        # Step 3: Commit the file to Shopify
-        file_info = commit_file_to_shopify(filename, location_url)
+        # Step 3: Commit the file to Shopify
+        file_info = commit_file_to_shopify(filename, location_url, alt_text)
         if file_info:
             # file_info contains the files array
             file_data = file_info[0]
             gid = file_data['id']
             # Get the URL
             url = None
             if 'url' in file_data:
                 url = file_data['url']
             elif 'image' in file_data and file_data['image']:
                 url = file_data['image'].get('url')
             if not url:
                 # Get URL via get_image_url_for_gid
                 url = get_image_url_for_gid(gid)
             
             return url, gid
         else:
             return None, None
 
     # Function to get all files from Shopify
-    def get_all_files():
-        all_files = {}
-        has_next_page = True
-        cursor = None
+    def get_all_files():
+        all_files = {}
+        has_next_page = True
+        cursor = None
 
         while has_next_page:
             query = f"""
             {{
             files(first: 250{' , after: "' + cursor + '"' if cursor else ''}) {{
                 edges {{
                 node {{
                     id
                     alt
                     ... on GenericFile {{
                     url
                     }}
                     ... on MediaImage {{
                     image {{
                         url
                     }}
                     }}
                 }}
                 cursor
                 }}
                 pageInfo {{
                 hasNextPage
                 }}
             }}
             }}
             """
             response = requests.post(GRAPHQL_URL, json={"query": query}, headers=graphql_headers)
             if response.status_code == 200:
                 data = response.json()
                 if "data" in data and "files" in data["data"]:
-                    for file in data["data"]["files"]["edges"]:
-                        node = file["node"]
-                        gid = node["id"]
-                        alt = node["alt"]
-                        url = None
-                        if 'url' in node:
-                            url = node['url']
-                        elif 'image' in node and node['image']:
-                            url = node['image'].get('url')
-                        if alt:
-                            all_files[alt] = (gid, url)
-                        cursor = file["cursor"]
+                    for file in data["data"]["files"]["edges"]:
+                        node = file["node"]
+                        gid = node["id"]
+                        alt = node["alt"]
+                        url = None
+                        if 'url' in node:
+                            url = node['url']
+                        elif 'image' in node and node['image']:
+                            url = node['image'].get('url')
+                        filename = ""
+                        if url:
+                            try:
+                                filename = os.path.basename(urllib.parse.urlparse(url).path)
+                            except Exception:
+                                filename = os.path.basename(url)
+
+                        if filename:
+                            all_files[filename] = (gid, url)
+                        if alt:
+                            all_files[alt] = (gid, url)
+                        cursor = file["cursor"]
 
                     has_next_page = data["data"]["files"]["pageInfo"]["hasNextPage"]
                 else:
                     print("No files found or error in response.")
                     return None
             else:
                 print(f"Error fetching files. Status code: {response.status_code}")
                 return None
 
-        return all_files
-
-    # Function to update or create metafields for a product
-    def update_metafields(handle, metafields, existing_files, row_index, df):
-        product_id = handle
-        if not product_id:
-            print(f"Skipping metafield update for missing product ID.")
-            return
+        return all_files
+
+    def parse_alt_values(raw_alt):
+        if raw_alt is None:
+            return []
+        if isinstance(raw_alt, float) and pd.isna(raw_alt):
+            return []
+        if isinstance(raw_alt, list):
+            return [str(item) if item is not None else '' for item in raw_alt]
+        if isinstance(raw_alt, str):
+            stripped = raw_alt.strip()
+            if not stripped:
+                return []
+            try:
+                parsed = json.loads(stripped)
+                if isinstance(parsed, list):
+                    return [str(item) if item is not None else '' for item in parsed]
+                if isinstance(parsed, str):
+                    return [parsed]
+            except json.JSONDecodeError:
+                return [item.strip() for item in stripped.split(',')]
+            return [stripped]
+        return [str(raw_alt)]
+
+    # Function to update or create metafields for a product
+    def update_metafields(handle, metafields, metafield_alts, existing_files, row_index, df):
+        product_id = handle
+        if not product_id:
+            print(f"Skipping metafield update for missing product ID.")
+            return
 
         # Fetch current metafields to handle deletion if necessary
         current_metafields_url = f"{BASE_URL}/products/{product_id}/metafields.json"
         response = requests.get(current_metafields_url, headers=headers)
         current_metafields = response.json().get('metafields', []) if response.status_code == 200 else []
         current_metafields_dict = {f"{mf['namespace']}.{mf['key']}": mf['id'] for mf in current_metafields}
 
-        for column, value in metafields.items():
-            key_type_str = column.replace('Metafield: ', '').split(' ')
-            key = key_type_str[0]
-            field_type = key_type_str[1].replace('[', '').replace(']', '')
-
-            namespace, key = key.split('.')
-
-            # Handle deletion if the value is None
-            if pd.isna(value) or value is None:
-                metafield_key = f"{namespace}.{key}"
-                if metafield_key in current_metafields_dict:
+        for column, value in metafields.items():
+            key_type_str = column.replace('Metafield: ', '').split(' ')
+            key = key_type_str[0]
+            field_type = key_type_str[1].replace('[', '').replace(']', '')
+
+            namespace, key = key.split('.')
+            alt_key = f"{namespace}.{key}"
+            alt_values = parse_alt_values(metafield_alts.get(alt_key) if metafield_alts else None)
+
+            # Handle deletion if the value is None
+            if pd.isna(value) or value is None:
+                metafield_key = f"{namespace}.{key}"
+                if metafield_key in current_metafields_dict:
                     delete_metafield(product_id, current_metafields_dict[metafield_key])
                 continue  # Skip to the next metafield if it's being deleted
 
             # For file_reference metafields
-            if field_type == 'file_reference':
-                if isinstance(value, str):
-                    if value.startswith('gid://'):
-                        # It's already a GID, use it directly
-                        value_gid = value
-                    elif value.startswith('http'):
-                        # It's a URL, need to find the GID
-                        filename = os.path.basename(value)
-                        if filename in existing_files:
-                            value_gid = existing_files[filename][0]
+            if field_type == 'file_reference':
+                if isinstance(value, str):
+                    if value.startswith('gid://'):
+                        # It's already a GID, use it directly
+                        value_gid = value
+                    elif value.startswith('http'):
+                        # It's a URL, need to find the GID
+                        filename = os.path.basename(value)
+                        if filename in existing_files:
+                            value_gid = existing_files[filename][0]
                         else:
                             value_gid = None
                     else:
                         # It's a filename
                         filename = value
                         if filename in existing_files:
                             gid, url = existing_files[filename]
                             value_gid = gid
                             # Replace cell value with GID
                             df.at[row_index, column] = value_gid
                         else:
                             # Upload the image
-                            file_path_local = os.path.join(IMAGE_FOLDER, filename)
-                            if os.path.exists(file_path_local):
-                                url, gid = upload_image_to_shopify(file_path_local)
-                                if gid:
-                                    existing_files[filename] = (gid, url)
-                                    value_gid = gid
-                                    # Replace cell value with GiID
-                                    df.at[row_index, column] = url
+                            file_path_local = os.path.join(IMAGE_FOLDER, filename)
+                            if os.path.exists(file_path_local):
+                                alt_text = alt_values[0] if alt_values else None
+                                url, gid = upload_image_to_shopify(file_path_local, alt_text)
+                                if gid:
+                                    existing_files[filename] = (gid, url)
+                                    value_gid = gid
+                                    # Replace cell value with GiID
+                                    df.at[row_index, column] = url
                                 else:
                                     print(f"Failed to upload image {filename}")
                                     continue
                             else:
                                 print(f"Image file {filename} not found in local folder.")
                                 continue
                 else:
                     # Value is not a string, skip
                     continue
 
                 if value_gid:
                     metafield_data = {
                         "metafield": {
                             "namespace": namespace,
                             "key": key,
                             "value": value_gid,
                             "type": field_type.strip()  # Strip any trailing spaces or newlines from type
                         }
                     }
                 else:
                     print(f"Cannot find GID for file {filename}")
                     continue
 
-            elif field_type == 'list.file_reference':
-                if isinstance(value, str):
-                    file_names = [filename.strip() for filename in value.split(",") if filename.strip()]
-                    file_gids = []  # Store all GIDs for the list
-
-                    print(f"🔍 Processing metafield '{key}' for product {product_id} with {len(file_names)} files: {file_names}")
-
-                    for filename in file_names:
-                        value_gid = None  # Reset for each file
-                        
-                        if filename.startswith('gid://'):
-                            value_gid = filename
-                            print(f"✅ Using existing GID: {value_gid}")
-
-                        elif filename.startswith('http'):
-                            base_filename = os.path.basename(filename)
-                            value_gid = existing_files.get(base_filename, [None])[0]
+            elif field_type == 'list.file_reference':
+                if isinstance(value, str):
+                    file_names = [filename.strip() for filename in value.split(",") if filename.strip()]
+                    file_gids = []  # Store all GIDs for the list
+
+                    print(f"🔍 Processing metafield '{key}' for product {product_id} with {len(file_names)} files: {file_names}")
+
+                    for idx, filename in enumerate(file_names):
+                        value_gid = None  # Reset for each file
+
+                        if filename.startswith('gid://'):
+                            value_gid = filename
+                            print(f"✅ Using existing GID: {value_gid}")
+
+                        elif filename.startswith('http'):
+                            base_filename = os.path.basename(filename)
+                            value_gid = existing_files.get(base_filename, [None])[0]
                             print(f"🔍 Retrieved GID from existing_files for {base_filename}: {value_gid}")
 
                         else:
                             # It's a filename
                             if filename in existing_files:
                                 gid, url = existing_files[filename]
                                 value_gid = gid
                                 df.at[row_index, column] = value_gid
 
                                 print(f"✅ Found existing upload for {filename}, GID: {value_gid}")
                             else:
                                 # Upload image (with compression fix)
-                                file_path_local = os.path.join(IMAGE_FOLDER, filename)
-                                if os.path.exists(file_path_local):
-                                    print(f"📤 Uploading {filename} to Shopify...")
-                                    url, gid = upload_image_to_shopify(file_path_local)  # Updated function call
-                                    if gid:
-                                        existing_files[filename] = (gid, url)
-                                        value_gid = gid
-                                        df.at[row_index, column] = url
-
+                                file_path_local = os.path.join(IMAGE_FOLDER, filename)
+                                if os.path.exists(file_path_local):
+                                    print(f"📤 Uploading {filename} to Shopify...")
+                                    alt_text = alt_values[idx] if idx < len(alt_values) else None
+                                    url, gid = upload_image_to_shopify(file_path_local, alt_text)  # Updated function call
+                                    if gid:
+                                        existing_files[filename] = (gid, url)
+                                        value_gid = gid
+                                        df.at[row_index, column] = url
+
                                         print(f"✅ Successfully uploaded {filename}, new GID: {value_gid}")
                                     else:
                                         print(f"❌ Failed to upload image {filename}")
                                         continue
                                 else:
                                     print(f"⚠️ Image file {filename} not found in local folder.")
                                     continue
 
                         if value_gid:
                             file_gids.append(value_gid)
                         else:
                             print(f"❌ Cannot find GID for file {filename}")
 
                     # ✅ Debugging: Show final GIDs list before updating Shopify
                     print(f"📝 Final GID list for metafield '{key}': {file_gids}")
 
                     # ✅ Only update Shopify if we have at least one valid GID
                     if file_gids:
                         metafield_data = {
                             "metafield": {
                                 "namespace": namespace,
                                 "key": key,
                                 "value":  json.dumps(file_gids),  # Store as a valid list
                                 "type": "list.file_reference"  # Ensure correct metafield type
                             }
diff --git a/Shopify Bulk Tool.py	 b/Shopify Bulk Tool.py	
index a227c98f97a01d759c62ab43087c8435f4dc01b9..9313edb90ff82dc43ff39a98099d8fed0b322eb0 100644
--- a/Shopify Bulk Tool.py	
+++ b/Shopify Bulk Tool.py	
@@ -1438,148 +1589,152 @@ def run_uploader_logic():
                         "value": value,
                         "type": field_type.strip()  # Strip any trailing spaces or newlines from type
                     }
                 }
                 print(f"Other metafield data: {metafield_data}")
 
 
             url = f"{BASE_URL}/products/{product_id}/metafields.json"
 
             print(f"📡 Sending request to Shopify API: {url}")
             print(f"🔍 Headers: {json.dumps(headers, indent=2)}")
             print(f"📝 Payload: {json.dumps(metafield_data, indent=2)}")
 
             response = requests.post(url, headers=headers, json=metafield_data)
 
             print(f"📩 Shopify Response: {response.status_code}")
 
             if response.status_code in [200, 201]:
                 print(f"✅ Successfully updated metafield {namespace}.{key} for product {product_id}")
             else:
                 print(f"❌ Failed to update metafield {namespace}.{key} for product {product_id}: {response.status_code}")
                 print(f"⚠️ Response Body: {response.text}")
 
 
 
-    def update_variant_metafields(variant_id, metafields, existing_files, row_index, df):
-        if not variant_id:
-            print("Skipping metafield update for missing variant ID.")
-            return
-
-        current_metafields_url = f"{BASE_URL}/variants/{variant_id}/metafields.json"
+    def update_variant_metafields(variant_id, metafields, metafield_alts, existing_files, row_index, df):
+        if not variant_id:
+            print("Skipping metafield update for missing variant ID.")
+            return
+
+        current_metafields_url = f"{BASE_URL}/variants/{variant_id}/metafields.json"
         response = requests.get(current_metafields_url, headers=headers)
         current_metafields = response.json().get('metafields', []) if response.status_code == 200 else []
         current_metafields_dict = {f"{mf['namespace']}.{mf['key']}": mf['id'] for mf in current_metafields}
 
-        for column, value in metafields.items():
-            key_type_str = column.replace('Variant Metafield: ', '').split(' ')
-            key = key_type_str[0]
-            field_type = key_type_str[1].replace('[', '').replace(']', '') if len(key_type_str) > 1 else 'single_line_text_field'
-
-            namespace, key = key.split('.')
-
-            if pd.isna(value) or value is None:
-                metafield_key = f"{namespace}.{key}"
-                if metafield_key in current_metafields_dict:
-                    delete_metafield(variant_id, current_metafields_dict[metafield_key])
+        for column, value in metafields.items():
+            key_type_str = column.replace('Variant Metafield: ', '').split(' ')
+            key = key_type_str[0]
+            field_type = key_type_str[1].replace('[', '').replace(']', '') if len(key_type_str) > 1 else 'single_line_text_field'
+
+            namespace, key = key.split('.')
+            alt_key = f"{namespace}.{key}"
+            alt_values = parse_alt_values(metafield_alts.get(alt_key) if metafield_alts else None)
+
+            if pd.isna(value) or value is None:
+                metafield_key = f"{namespace}.{key}"
+                if metafield_key in current_metafields_dict:
+                    delete_metafield(variant_id, current_metafields_dict[metafield_key])
                 continue
 
-            if field_type == 'file_reference':
-                if isinstance(value, str):
-                    if value.startswith('gid://'):
-                        value_gid = value
-                    elif value.startswith('http'):
-                        filename = os.path.basename(value)
-                        value_gid = existing_files.get(filename, (None,))[0]
-                    else:
-                        filename = value
-                        if filename in existing_files:
-                            gid, url = existing_files[filename]
-                            value_gid = gid
-                            df.at[row_index, column] = value_gid
-                        else:
-                            file_path_local = os.path.join(IMAGE_FOLDER, filename)
-                            if os.path.exists(file_path_local):
-                                url, gid = upload_image_to_shopify(file_path_local)
-                                if gid:
-                                    existing_files[filename] = (gid, url)
-                                    value_gid = gid
-                                    df.at[row_index, column] = url
-                                else:
-                                    print(f"Failed to upload image {filename}")
-                                    continue
-                            else:
-                                print(f"Image file {filename} not found in local folder.")
-                                continue
+            if field_type == 'file_reference':
+                if isinstance(value, str):
+                    if value.startswith('gid://'):
+                        value_gid = value
+                    elif value.startswith('http'):
+                        filename = os.path.basename(value)
+                        value_gid = existing_files.get(filename, (None,))[0]
+                    else:
+                        filename = value
+                        if filename in existing_files:
+                            gid, url = existing_files[filename]
+                            value_gid = gid
+                            df.at[row_index, column] = value_gid
+                        else:
+                            file_path_local = os.path.join(IMAGE_FOLDER, filename)
+                            if os.path.exists(file_path_local):
+                                alt_text = alt_values[0] if alt_values else None
+                                url, gid = upload_image_to_shopify(file_path_local, alt_text)
+                                if gid:
+                                    existing_files[filename] = (gid, url)
+                                    value_gid = gid
+                                    df.at[row_index, column] = url
+                                else:
+                                    print(f"Failed to upload image {filename}")
+                                    continue
+                            else:
+                                print(f"Image file {filename} not found in local folder.")
+                                continue
                 else:
                     continue
 
                 if value_gid:
                     metafield_data = {
                         "metafield": {
                             "namespace": namespace,
                             "key": key,
                             "value": value_gid,
                             "type": field_type.strip()
                         }
                     }
                 else:
                     print(f"Cannot find GID for file {filename}")
                     continue
 
-            elif field_type == 'list.file_reference':
-                if isinstance(value, str):
-                    file_names = [filename.strip() for filename in value.split(",") if filename.strip()]
-                    file_gids = []
-
-                    for filename in file_names:
-                        value_gid = None
-
-                        if filename.startswith('gid://'):
-                            value_gid = filename
-                        elif filename.startswith('http'):
-                            base_filename = os.path.basename(filename)
-                            value_gid = existing_files.get(base_filename, [None])[0]
-                        else:
-                            if filename in existing_files:
-                                gid, url = existing_files[filename]
-                                value_gid = gid
-                                df.at[row_index, column] = value_gid
-                            else:
-                                file_path_local = os.path.join(IMAGE_FOLDER, filename)
-                                if os.path.exists(file_path_local):
-                                    url, gid = upload_image_to_shopify(file_path_local)
-                                    if gid:
-                                        existing_files[filename] = (gid, url)
-                                        value_gid = gid
-                                        df.at[row_index, column] = url
-                                    else:
-                                        print(f"Failed to upload image {filename}")
-                                        continue
-                                else:
-                                    print(f"Image file {filename} not found in local folder.")
-                                    continue
+            elif field_type == 'list.file_reference':
+                if isinstance(value, str):
+                    file_names = [filename.strip() for filename in value.split(",") if filename.strip()]
+                    file_gids = []
+
+                    for idx, filename in enumerate(file_names):
+                        value_gid = None
+
+                        if filename.startswith('gid://'):
+                            value_gid = filename
+                        elif filename.startswith('http'):
+                            base_filename = os.path.basename(filename)
+                            value_gid = existing_files.get(base_filename, [None])[0]
+                        else:
+                            if filename in existing_files:
+                                gid, url = existing_files[filename]
+                                value_gid = gid
+                                df.at[row_index, column] = value_gid
+                            else:
+                                file_path_local = os.path.join(IMAGE_FOLDER, filename)
+                                if os.path.exists(file_path_local):
+                                    alt_text = alt_values[idx] if idx < len(alt_values) else None
+                                    url, gid = upload_image_to_shopify(file_path_local, alt_text)
+                                    if gid:
+                                        existing_files[filename] = (gid, url)
+                                        value_gid = gid
+                                        df.at[row_index, column] = url
+                                    else:
+                                        print(f"Failed to upload image {filename}")
+                                        continue
+                                else:
+                                    print(f"Image file {filename} not found in local folder.")
+                                    continue
 
                         if value_gid:
                             file_gids.append(value_gid)
                         else:
                             print(f"Cannot find GID for file {filename}")
 
                     if file_gids:
                         metafield_data = {
                             "metafield": {
                                 "namespace": namespace,
                                 "key": key,
                                 "value": json.dumps(file_gids),
                                 "type": "list.file_reference"
                             }
                         }
                     else:
                         print(f"Skipping metafield update for '{key}' because the file list is empty.")
                         continue
                 else:
                     print(f"Skipping non-string value for metafield {key}.")
                     continue
 
             else:
                 if field_type == 'rich_text_field':
                     try:
diff --git a/Shopify Bulk Tool.py	 b/Shopify Bulk Tool.py	
index a227c98f97a01d759c62ab43087c8435f4dc01b9..9313edb90ff82dc43ff39a98099d8fed0b322eb0 100644
--- a/Shopify Bulk Tool.py	
+++ b/Shopify Bulk Tool.py	
@@ -2076,81 +2231,89 @@ def run_uploader_logic():
                             else:
                                 print(f"Image file {filename} not found in local folder.")
                     else:
                         continue  # Skip if image URL is not valid
 
                 # Update product images with alt texts
                 if images:
                     update_product_images(product_id, images, alt_texts)
 
                     # Retrieve the updated images to get their IDs
                     url = f"{BASE_URL}/products/{product_id}/images.json"
                     response = requests.get(url, headers=headers)
                     if response.status_code == 200:
                         product_images = response.json().get('images', [])
                         for image_data in product_images:
                             image_src = image_data.get('src')
                             image_id = image_data.get('id')
                             if image_src in images:
                                 index_in_list = images.index(image_src)
                                 alt_text = alt_texts[index_in_list]
                                 
                     else:
                         print(f"Failed to retrieve images for product {product_id}: {response.status_code}, {response.text}")
 
  # Process metafields
-                metafields = {}
-                variant_metafields = {}
-                for column in df.columns:
-                    if column.startswith('Metafield:'):
-                        value = row[column]
-                        if '[rich_text_field]' in column and isinstance(value, str) and '<' in value and '>' in value:
-                            try:
-                                json_value = html_to_shopify_json(value)
-                                metafields[column] = json_value
-                            except Exception as e:
-                                print(f"Error parsing HTML for {column}: {e}")
-                        else:
-                            metafields[column] = value
-                    elif column.startswith('Variant Metafield:'):
-                        value = row[column]
-                        if '[rich_text_field]' in column and isinstance(value, str) and '<' in value and '>' in value:
-                            try:
-                                json_value = html_to_shopify_json(value)
-                                variant_metafields[column] = json_value
-                            except Exception as e:
-                                print(f"Error parsing HTML for {column}: {e}")
-                        else:
-                            variant_metafields[column] = value
-
-
-                # Update metafields for the product
-                if metafields:
-                    if pd.notna(product_id):
-                        update_metafields(product_id, metafields, existing_files, index, df)
-                    else:
-                        print(f"Product ID missing for row {index}, cannot update metafields.")
+                metafields = {}
+                variant_metafields = {}
+                metafield_alts = {}
+                variant_metafield_alts = {}
+                for column in df.columns:
+                    if column.startswith('Metafield Alt:'):
+                        alt_key = column.replace('Metafield Alt: ', '').strip()
+                        metafield_alts[alt_key] = row[column]
+                    elif column.startswith('Variant Metafield Alt:'):
+                        alt_key = column.replace('Variant Metafield Alt: ', '').strip()
+                        variant_metafield_alts[alt_key] = row[column]
+                    elif column.startswith('Metafield:'):
+                        value = row[column]
+                        if '[rich_text_field]' in column and isinstance(value, str) and '<' in value and '>' in value:
+                            try:
+                                json_value = html_to_shopify_json(value)
+                                metafields[column] = json_value
+                            except Exception as e:
+                                print(f"Error parsing HTML for {column}: {e}")
+                        else:
+                            metafields[column] = value
+                    elif column.startswith('Variant Metafield:'):
+                        value = row[column]
+                        if '[rich_text_field]' in column and isinstance(value, str) and '<' in value and '>' in value:
+                            try:
+                                json_value = html_to_shopify_json(value)
+                                variant_metafields[column] = json_value
+                            except Exception as e:
+                                print(f"Error parsing HTML for {column}: {e}")
+                        else:
+                            variant_metafields[column] = value
+
+
+                # Update metafields for the product
+                if metafields:
+                    if pd.notna(product_id):
+                        update_metafields(product_id, metafields, metafield_alts, existing_files, index, df)
+                    else:
+                        print(f"Product ID missing for row {index}, cannot update metafields.")
 
             # Prepare the variant update data if there is a variant
              # Prepare the variant update data if there is a variant
             if variant_name:
 
 
                 variant_data = {
                 "variant": {
                     "id": variant_id,
                     "price": row['Variant Price'],
                     "option1": row.get('Option1 Value', ""),  # Add Option1 Value
                     "option2": row.get('Option2 Value', ""),  # Add Option2 Value
                     "option3": row.get('Option3 Value', "")   # Add Option3 Value
                 }
             }   
 
             # Conditionally add 'sku' if present and valid
             if row.get('Variant SKU'):
                 variant_data["variant"]["sku"] = row['Variant SKU']
 
             # Conditionally add 'barcode' if present and valid
             if row.get('Variant Barcode'):
                 variant_data["variant"]["barcode"] = row['Variant Barcode']
 
             # Conditionally add 'weight' if present and valid
diff --git a/Shopify Bulk Tool.py	 b/Shopify Bulk Tool.py	
index a227c98f97a01d759c62ab43087c8435f4dc01b9..9313edb90ff82dc43ff39a98099d8fed0b322eb0 100644
--- a/Shopify Bulk Tool.py	
+++ b/Shopify Bulk Tool.py	
@@ -2189,52 +2352,52 @@ def run_uploader_logic():
 
                 # Handle market-specific pricing dynamically
                 for column, market_name in pricing_columns.items():
                     print(f"Processing column '{column}' for market '{market_name}'...")
 
                     if market_name in market_names:
                         print(f"Market '{market_name}' found in Shopify markets.")
                         price_amount = row.get(column)
                         print(f"Price from column '{column}': {price_amount}")
 
                         price_list_id = get_price_list_id_for_market(market_name)
                         print(f"Price list ID for market '{market_name}': {price_list_id}")
 
                         if price_list_id and price_amount:
                             print(f"Adding fixed price for market '{market_name}', variant ID '{variant_id}', price '{price_amount}'")
                             add_fixed_price_for_market(price_list_id, variant_id, price_amount)
                             print(f"Fixed price added successfully for market '{market_name}'.")
                         else:
                             if not price_list_id:
                                 print(f"[WARNING] Price list ID is missing for market '{market_name}'. Unable to add fixed price.")
                             if not price_amount:
                                 print(f"[WARNING] Price amount is missing in column '{column}' for row {index}. Skipping this market.")
                     else:
                         print(f"[INFO] Market '{market_name}' from column '{column}' not found in Shopify markets. Skipping.")
 
-                        if variant_metafields:
-                            update_variant_metafields(variant_id, variant_metafields, existing_files, index, df)
+                        if variant_metafields:
+                            update_variant_metafields(variant_id, variant_metafields, variant_metafield_alts, existing_files, index, df)
 
                 else:
                     print(f"Failed to process variant for handle '{handle}'.")
 
 
 
 
         # After processing all rows, save the updated DataFrame back to Excel
         df.to_excel(file_path, index=False)
         print("Spreadsheet updated with new image URLs.")
 
     # Function to prompt the user to select a file
     def get_file_path():
         root = tk.Tk()
         root.withdraw()  # Hide the main window
         file_path = filedialog.askopenfilename(
             title="Select Excel File",
             filetypes=[("Excel files", "*.xlsx")]
         )
         return file_path
 
     # Example usage
     file_path = get_file_path()
     if file_path:
         upload_changes_from_spreadsheet(file_path)
diff --git a/Shopify Bulk Tool.py	 b/Shopify Bulk Tool.py	
index a227c98f97a01d759c62ab43087c8435f4dc01b9..9313edb90ff82dc43ff39a98099d8fed0b322eb0 100644
--- a/Shopify Bulk Tool.py	
+++ b/Shopify Bulk Tool.py	
@@ -2488,160 +2651,160 @@ def collection_run_uploader_logic():
         data = response.json()
         if "data" in data and "stagedUploadsCreate" in data["data"]:
             return data["data"]["stagedUploadsCreate"]["stagedTargets"][0]
         else:
             print("Error in staged upload creation:", data)
             return None
 
     # Function to upload the file to the staging URL
     def upload_file_to_staging(staging_target, file_path):
         url = staging_target["url"]
         files = {"file": open(file_path, "rb")}
         form_data = {param["name"]: param["value"] for param in staging_target["parameters"]}
         response = requests.post(url, data=form_data, files=files)
 
         # Parsing the XML response to get the file location
         if response.status_code == 201:
             print(f"File {file_path} successfully uploaded to staging URL.")
             xml_response = ET.fromstring(response.text)
             location_url = xml_response.find('Location').text  # Extracting Location URL from XML
             return location_url
         else:
             print(f"Failed to upload file. Status: {response.status_code}, {response.text}")
             return None
 
     # Function to commit the file with fileCreate
-    def commit_file_to_shopify(file_name, original_source):
-        query = """
-        mutation fileCreate($files: [FileCreateInput!]!) {
-        fileCreate(files: $files) {
-            files {
-            id  # Fetch the gid after file creation
+    def commit_file_to_shopify(file_name, original_source, alt_text=None):
+        query = """
+        mutation fileCreate($files: [FileCreateInput!]!) {
+        fileCreate(files: $files) {
+            files {
+            id  # Fetch the gid after file creation
             alt
             createdAt
             ... on GenericFile {
                 url
             }
             ... on MediaImage {
                 image {
                 url
                 }
             }
             }
             userErrors {
             code
             field
             message
             }
         }
-        }
-        """
-        variables = {
-            "files": [
-                {
-                    "alt": file_name,
-                    "originalSource": original_source
-                }
-            ]
-        }
+        }
+        """
+        variables = {
+            "files": [
+                {
+                    "alt": alt_text if alt_text is not None else file_name,
+                    "originalSource": original_source
+                }
+            ]
+        }
         response = requests.post(GRAPHQL_URL, json={"query": query, "variables": variables}, headers=graphql_headers)
         data = response.json()
         if "data" in data and "fileCreate" in data["data"]:
             file_info = data["data"]["fileCreate"]["files"]
             gid = file_info[0]["id"]
             print(f"File {file_name} successfully committed to Shopify.")
             print(f"GID: {gid}")  # Print the gid here
             return file_info
         else:
             print("Error in file commit:", data)
             return None
 
     # Function to get image URL for a given GID
     def get_image_url_for_gid(gid):
         query = f"""
         {{
         node(id: "{gid}") {{
             ... on GenericFile {{
             url
             }}
             ... on MediaImage {{
             image {{
                 url
             }}
             }}
         }}
         }}
         """
         response = requests.post(GRAPHQL_URL, json={"query": query}, headers=graphql_headers)
         if response.status_code == 200:
             data = response.json()
             node = data.get('data', {}).get('node', {})
             url = None
             if 'url' in node:
                 url = node['url']
             elif 'image' in node and node['image']:
                 url = node['image'].get('url')
             return url
         else:
             print(f"Failed to get image URL for GID {gid}")
             return None
 
-    def upload_image_to_shopify(file_path):
-        filename = os.path.basename(file_path)
-
-        # Normalize the filename before accessing it
-        normalized_filename = normalize_filename(filename)
-        folder_path = os.path.dirname(file_path)
-
-        file_path = os.path.join(os.path.dirname(file_path), normalized_filename)
+    def upload_image_to_shopify(file_path, alt_text=None):
+        filename = os.path.basename(file_path)
+
+        # Normalize the filename before accessing it
+        normalized_filename = normalize_filename(filename)
+        folder_path = os.path.dirname(file_path)
+
+        file_path = os.path.join(os.path.dirname(file_path), normalized_filename)
 
         
         # Check if file exists after normalization
         if not os.path.exists(file_path):
             print(f"Image file {normalized_filename} not found in local folder.")
             return None, None
 
 
           # Ensure the filename is URL-encoded
         encoded_filename = encode_filename(filename)    
     
     
         # Extract the file name and MIME type
         mime_type = "image/jpeg" if filename.lower().endswith(".jpg") else "image/png"
 
         # Step 1: Create the staged upload
         staging_target = staged_upload_create(filename, mime_type)
         if not staging_target:
             return None, None
 
         # Step 2: Upload the file to the staging URL
-        location_url = upload_file_to_staging(staging_target, file_path)
-        if not location_url:
-            return None, None
-
-        # Step 3: Commit the file to Shopify
-        file_info = commit_file_to_shopify(filename, location_url)
+        location_url = upload_file_to_staging(staging_target, file_path)
+        if not location_url:
+            return None, None
+
+        # Step 3: Commit the file to Shopify
+        file_info = commit_file_to_shopify(filename, location_url, alt_text)
         if file_info:
             # file_info contains the files array
             file_data = file_info[0]
             gid = file_data['id']
             # Get the URL
             url = None
             if 'url' in file_data:
                 url = file_data['url']
             elif 'image' in file_data and file_data['image']:
                 url = file_data['image'].get('url')
             if not url:
                 # Get URL via get_image_url_for_gid
                 url = get_image_url_for_gid(gid)
                 
             print(f"✅ Image uploaded successfully: GID={gid}, URL={url}")
 
             return url, gid
         else:
             return None, None
 
     def get_all_files():
         """Retrieves all existing file references in Shopify."""
         print("📡 Fetching all existing Shopify files...")
         all_files = {}
         has_next_page = True
diff --git a/Shopify Bulk Tool.py	 b/Shopify Bulk Tool.py	
index a227c98f97a01d759c62ab43087c8435f4dc01b9..9313edb90ff82dc43ff39a98099d8fed0b322eb0 100644
--- a/Shopify Bulk Tool.py	
+++ b/Shopify Bulk Tool.py	
@@ -2653,58 +2816,67 @@ def collection_run_uploader_logic():
             files(first: 250{' , after: "' + cursor + '"' if cursor else ''}) {{
                 edges {{
                 node {{
                     id
                     alt
                     ... on GenericFile {{
                     url
                     }}
                     ... on MediaImage {{
                     image {{
                         url
                     }}
                     }}
                 }}
                 cursor
                 }}
                 pageInfo {{
                 hasNextPage
                 }}
             }}
             }}
             """
             response = requests.post(GRAPHQL_URL, json={"query": query}, headers=headers)
             if response.status_code == 200:
                 data = response.json()
-                for file in data.get("data", {}).get("files", {}).get("edges", []):
-                    node = file["node"]
-                    gid = node["id"]
-                    alt = node["alt"]
-                    url = node.get('url') or node.get('image', {}).get('url')
-                    if alt:
-                        all_files[alt] = (gid, url)
-                    cursor = file["cursor"]
+                for file in data.get("data", {}).get("files", {}).get("edges", []):
+                    node = file["node"]
+                    gid = node["id"]
+                    alt = node["alt"]
+                    url = node.get('url') or node.get('image', {}).get('url')
+                    filename = ""
+                    if url:
+                        try:
+                            filename = os.path.basename(urllib.parse.urlparse(url).path)
+                        except Exception:
+                            filename = os.path.basename(url)
+
+                    if filename:
+                        all_files[filename] = (gid, url)
+                    if alt:
+                        all_files[alt] = (gid, url)
+                    cursor = file["cursor"]
 
                 has_next_page = data["data"]["files"]["pageInfo"]["hasNextPage"]
                 print(f"✅ {len(all_files)} files retrieved so far...")
             else:
                 print(f"❌ Error fetching files. Response: {response.status_code}")
                 return None
 
         print(f"✅ Finished fetching {len(all_files)} existing files.")
         return all_files
 
     def update_metafields(collection_id, metafields, existing_files, row_index, df):
         """Handles metafield updates for collections, including file uploads and deletions."""
         
         if not collection_id:
             print(f"⚠️ Skipping metafield update: Missing collection ID.")
             return
 
         print(f"\n📡 Fetching existing metafields for Collection ID: {collection_id}...")
         
         # Fetch current metafields
         url = f"{BASE_URL}/collections/{collection_id}/metafields.json"
         response = requests.get(url, headers=headers)
         current_metafields = response.json().get('metafields', []) if response.status_code == 200 else []
         current_metafields_dict = {f"{mf['namespace']}.{mf['key']}": mf['id'] for mf in current_metafields}
 
 
EOF
)
