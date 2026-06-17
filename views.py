from gettext import GNUTranslations
import json
from operator import index
from optparse import Values
import os
import re
from csv import excel
import asyncio
import stat
from textwrap import indent
import aiofiles
from numpy import int64
import pandas as pd
from datetime import datetime
from pandas._libs import missing
from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response
from mcom_website.settings import MEDIA_ROOT, MEDIA_URL
from gpl_audit_tool_V1_1.extractors.command_extractor import CommandExtractor
from gpl_audit_tool_V1_1.extractors.table_extractor import TableExtractor
from gpl_audit_tool_V1_1.thread_pool import ThreadPool
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from collections import defaultdict
import zipfile
from asgiref.sync import sync_to_async
from rest_framework.decorators import api_view
from rest_framework.response import Response

from functools import partial

from concurrent.futures import ThreadPoolExecutor

from django.views.decorators.csrf import csrf_exempt

import asyncio
from django.conf import settings
from gpl_audit_tool_V1_1.process_scripts import (
    process_correction_script_generation,
    create_freqency_relation_script,
)
from gpl_audit_tool_V1_1.GPL_FREQ_REL_SCRIPTS import (
    EUtranFrequency_Definition,
    GeranFrequency_defination,
    Eutran_Freq_relation_creation_script,
    Eutran_freq_cell_relation_defination,
)


def format_excel_sheet(writer, sheet_name, df, startrow=0, startcol=0):
    """Apply formatting to an Excel sheet with adjustable start positions."""
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    header_format = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#000957",
            "border": 2,
            "font_color": "#ffffff",
            "align": "center",
            "valign": "vcenter",
        }
    )
    center_format = workbook.add_format(
        {
            "align": "center",
            "valign": "center",
            "border": 1,
            "border_color": "#000000",
            "bold": True,
        }
    )
    ok_format = workbook.add_format(
        {
            "bg_color": "#90EE90",
            "font_color": "#000000",
            "align": "center",
            "valign": "center",
        }
    )
    not_ok_format = workbook.add_format(
        {
            "bg_color": "#FF0000",
            "font_color": "#FFFFFF",
            "align": "center",
            "valign": "center",
        }
    )

    worksheet.set_row(startrow, 23)

    for col_num, col_name in enumerate(df.columns):
        worksheet.write(startrow, startcol + col_num, str(col_name), header_format)

        column_series = df[col_name]
        if isinstance(column_series, pd.DataFrame):
            column_series = column_series.iloc[:, 0]

        max_length = max(
            column_series.astype(str).fillna("").str.len().max(skipna=True) or 0,
            len(str(col_name)),
        )
        max_length = min(max_length, 255)
        worksheet.set_column(startcol + col_num, startcol + col_num, max_length + 5)

    for row_num in range(len(df)):
        worksheet.set_row(startrow + row_num + 1, 15)

        for col_num in range(len(df.columns)):
            cell_value = str(df.iloc[row_num, col_num])
            format_style = center_format
            if cell_value == "OK":
                format_style = ok_format
            elif cell_value == "NOT OK":
                format_style = not_ok_format
            elif cell_value == "Missing" or cell_value == "Missing in Post":
                format_style = workbook.add_format(
                    {
                        "bg_color": "#FF6347",
                        "font_color": "#FFFFFF",
                        "align": "center",
                        "valign": "center",
                    }
                )
            elif cell_value == "NA":
                format_style = workbook.add_format(
                    {
                        "bg_color": "#FCF259",
                        "font_color": "#222831",
                        "align": "center",
                        "valign": "center",
                    }
                )
            elif "|" in cell_value:
                format_style = workbook.add_format(
                    {
                        "font_color": "#FF0000",
                        "align": "center",
                        "valign": "center",
                        "bold": True,
                        "border": 1,
                        "border_color": "#000000",
                    }
                )

            worksheet.write(
                startrow + row_num + 1, startcol + col_num, cell_value, format_style
            )


# def get_commands(command_lines):
#     pool = ThreadPool()
#     result = {}

#     tasks = [
#         (
#             r"#+Start:\senbinfo\sAudit\s#+",
#             r"#+\sEND\sof\senbinfo\sAudit\s#+",
#             "enbinfo",
#         ),
#         (r"#+Start:\sCELL\sDATA\s#+", r"#+\sEND\sOF\sCELL\sDATA\s#+", "cell_data"),
#         (
#             r"#+Start:\sLTE\sGPL\sAudit\s#+",
#             r"#####END of LTE GPL Audit #######",
#             "gpl-para",
#         ),
#         (
#             r"#+Start:\sFeatureState\sAudit\s#+",
#             r"#+\sEND\sof\sFeatureState\sAudit\s#+",
#             "FeatureState",
#         ),
#         (
#             r"#+Start:\sEutranfrequency\sAudit\s#+",
#             r"#+\sEND\sof\sEutranfrequency\sAudit\s#+",
#             "Eutranfrequency",
#         ),
#         (
#             r"#+Start:\sEutranfreqRelation\sAudit\s#+",
#             r"#+\sEND\sof\sEutranfreqRelation\sAudit\s#+",
#             "EutranfreqRelation",
#         ),
#         (
#             r"#####Start:\s*CellRelation\s+Audit\s*#+",
#             r"#####\s*END\s+of\s+CellRelation\s+Audit\s*#+",
#             "CellRelation",
#         ),
#         (
#             r"#####Start:\sGeranFreqRelation\sAudit\s#######",
#             r"#####\sEND\sof\sGeranFreqRelation\sAudit\s#######",
#             "GeranFreqRelation",
#         ),
#     ]

#     for start, end, key in tasks:
#         pool.submit_task(
#             CommandExtractor.extract_command, command_lines, start, end, result, key
#         )

#     pool.wait_for_completion()

#     return result



def get_commands(command_lines):
    tasks = [
        (r"#+Start:\s*enbinfo\s*Audit\s*#+", r"#+\s*END\s+of\s+enbinfo\s*Audit\s*#+", "enbinfo"),
        (r"#+Start:\s*CELL\s*DATA\s*#+", r"#+\s*END\s+OF\s+CELL\s+DATA\s*#+", "cell_data"),
        (r"#+Start:\s*LTE\s*GPL\s*Audit\s*#+", r"#+\s*END\s+of\s+LTE\s+GPL\s+Audit\s*#+", "gpl-para"),
        (r"#+Start:\s*FeatureState\s*Audit\s*#+", r"#+\s*END\s+of\s+FeatureState\s*Audit\s*#+", "FeatureState"),
        (r"#+Start:\s*Eutranfrequency\s*Audit\s*#+", r"#+\s*END\s+of\s+Eutranfrequency\s*Audit\s*#+", "Eutranfrequency"),
        (r"#+Start:\s*EutranfreqRelation\s*Audit\s*#+", r"#+\s*END\s+of\s+EutranfreqRelation\s*Audit\s*#+", "EutranfreqRelation"),
        (r"#####Start:\s*CellRelation\s+Audit\s*#+", r"#####\s*END\s+of\s+CellRelation\s+Audit\s*#+", "CellRelation"),
        (r"#####Start:\s*GeranFreqRelation\s*Audit\s*#+", r"#####\s*END\s+of\s+GeranFreqRelation\s*Audit\s*#+", "GeranFreqRelation"),
    ]

    result = {}

    # ThreadPoolExecutor handles thread cleanup automatically
    with ThreadPoolExecutor() as executor:
        futures = {
            executor.submit(CommandExtractor.extract_command, command_lines, start, end): key
            for start, end, key in tasks
        }

        for future in futures:
            key = futures[future]
            result[key] = future.result()

    return result



def get_LTE_NR_GPL_AUDIT_commands(command_lines):
    pool = ThreadPool()
    result = {}

    tasks = [
        ############################################## LTE CELL DATA ########################################
        (
            r"#+Start:\s*LTE\s*CELL\s*DATA\s*#+",
            r"#+\s*END\s+OF\s+LTE\s+CELL\s+DATA\s*#+",
            "LTE_CELL_DATA",
        ),
        #################################################### NR CELL DATA ###################################
        (
            r"#+Start:\s*NR\s*CELL\s*DATA\s*#+",
            r"#+\s*END\s+OF\s+NR\s+CELL\s+DATA\s*#+",
            "NR_CELL_DATA",
        ),
        #################################################### NR CELL DATA ###################################
        (
            r"#+Start:\s*NR\s*GENBID\s*DATA\s*#+",
            r"#+\s*END\s+OF\s+NR\s+GNBID\s+DATA\s*#+",
            "NR_GNBID_DATA",
        ),
        ######################################################### LTE GPL Audit ###############################
        (
            r"#+Start:\s*LTE\s*GPL\s*Audit\s*#+",
            r"#+\s*END\s+of\s+LTE\s+GPL\s+Audit\s*#+",
            "LTE_GPL_Audit",
        ),
        ################################################### Relation Audit ####################################
        (
            r"#+Start:\s*Relation\s*Audit\s*#+",
            r"#+\s*END\s+of\s+Relation\s*Audit\s*#+",
            "Relation_Audit",
        ),
        ################################################ NR Baseline Audit #######################################
        (
            r"#+\s*NR\s*Baseline\s*Audit\s*#+",
            r"#+\s*END\s+of\s+NR\s+Baseline\s+Audit\s*#+",
            "NR_Baseline_Audit",
        ),
    ]

    for start, end, key in tasks:
        pool.submit_task(
            CommandExtractor.extract_command, command_lines, start, end, result, key
        )

    pool.wait_for_completion()
    return result


def process_files(file, log_path, result):
    """Process single log file using the new streaming TableExtractor."""
    
    # Flatten all required commands from the result dictionary
    target_commands = [cmd for cmd_list in result.values() for cmd in cmd_list]
    
    # Pass the log_path directly. NO lists, NO generators.
    extractor = TableExtractor(log_path, target_commands=target_commands)
    
    # Lte_nR_extract_tables remains completely untouched
    dfs = extract_tables(extractor, result)
    
    return file, dfs


def extract_tables(extractor, result):
    dataframes = {}
    for key, commands in result.items():
        # if key == "enbinfo":
        print("KEY:-> ", key)
        dataframes[key] = {}
        temp_dfs = []
        for command in commands:
            table = extractor.extract_table(command)
            print("table for command ", command, " is ", table)
            node_id = extractor.get_nodeID() or "UNKNOWN_NODE"
            if table:
                headers = [col.strip() for col in table[0].split(";")]
                max_cols = len(headers)
                data = [str(row).split(";") for row in table[1:]]
                data = [[i.strip() for i in val] for val in data]
                data = [
                    (
                        row + [""] * (max_cols - len(row))
                        if len(row) < max_cols
                        else row[:max_cols]
                    )
                    for row in data
                ]

                df = pd.DataFrame(data, columns=headers)

###################### new change - important Qci para not extracting without this section ###################################
                # if key == "EutranfreqRelation":
                #     print("\nHEADERS FOR EUTRANFREQRELATION:")
                #     print(headers)
                #     print("\n===== RAW EUTRANFREQRELATION DF =====")
                #     print(df.columns.tolist())

                #     qci_cols = [
                #         "a5Thr1RsrpFreqQciOffset",
                #         "a5Thr2RsrpFreqQciOffset",
                #         "qciProfileRef",
                #     ]

                #     existing_cols = [c for c in qci_cols if c in df.columns]

                #     print("Existing columns:", existing_cols)

                #     if existing_cols:
                #         print(df[existing_cols].head(20))
                #     print("\n===== POST QCI PARAMETERS CHECK =====")
                #     for col in existing_cols:
                #         print("Column:", col)
                #         print("Values:")
                #         print(df[col].head(20))
                #     print("========================================")
                    
###################### new change - important Qci para not extracting without this section ###################################

                if key == "gpl-para":
                    df = df.melt(
                        id_vars=["MO"] if "MO" in df.columns else [],
                        var_name="Parameter",
                        value_name="Value",
                    )

                    df["Value"] = (
                        df["Value"]
                        .apply(lambda x: str(x).split(" ")[0] if " " in x else x)
                        .astype(str)
                    )

                df.insert(0, "Node_ID", [node_id] * len(df))
                temp_dfs.append(df)

        if temp_dfs:
            merged_df = pd.concat(temp_dfs, axis=0, ignore_index=False)
            if key == "cell_data":
                merged_df = pd.concat(
                    temp_dfs,
                    axis=1,
                    ignore_index=False,
                    keys=["left", "right"],
                )

                merged_df.columns = merged_df.columns.map(
                    lambda x: f"{x[0]}_{x[1]}" if isinstance(x, tuple) else x
                )
                # print(merged_df.columns)
                merged_df.drop(columns=["right_Node_ID"], inplace=True)
            dataframes[key] = merged_df
    return dataframes


@api_view(["POST"])
def get_log_parser(request):
    log_files = request.FILES.getlist("log_files")
    baseURL = os.path.join(settings.MEDIA_ROOT, "4G_5G_GPL_AUDIT")
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")

    ################################################################## Create datetime-based folder ###########################################
    output_folder = os.path.join(
        baseURL, "parsed_dumps", f"parsed_files_{current_time}"
    )
    os.makedirs(output_folder, exist_ok=True, mode=0o777)

    #################################################################### Load command list ######################################################
    command_file_path = os.path.join(
        baseURL, "command_file", "GPL_Audit_command_updated.txt"
    )
    with open(command_file_path, "r") as command_file:
        command_lines = [line.strip() for line in command_file.readlines()]
    results = get_commands(command_lines)

    ############################################################# Save log files to disk ########################################################
    log_input_folder = os.path.join(baseURL, f"log_input_control/logs_{current_time}")
    os.makedirs(log_input_folder, exist_ok=True, mode=0o777)

    saved_files = []
    for file in log_files:
        file_path = os.path.join(log_input_folder, file.name)
        with open(file_path, "wb+") as destination:
            for chunk in file.chunks():
                destination.write(chunk)
        saved_files.append(file_path)

    ################################################# Extract node names from files #################################################################
    node_list = []
    for file in saved_files:
        if not (file.endswith(".log") or file.endswith(".txt")):
            return Response(
                {"error": "All files must be .log or .txt files."},
                status=status.HTTP_400_BAD_REQUEST,
            )

        with open(file, "r") as f:
            content = f.read()

        match = re.search(r"ManagedElement=([\w-]+)", content)
        if match:
            node = match.group(1)
            node_list.append({"node": node, "file_path": file})
        else:
            print("No node match found in:", file)

    all_download_urls = []

    ################################################### Process each node separately ###########################################################
    for item in node_list:
        node = item["node"]
        file_path = item["file_path"]

        all_dfs = defaultdict(list)
        all_dfs["Summary"] = []

        # Process only the current node file
        file, df_dict = process_files(file_path, file_path, results)

        for key, df in df_dict.items():
            all_dfs[key].append(df)

        excel_path = os.path.join(
            output_folder, f"{node}_GPL_PARSED_DUMP_{current_time}.xlsx"
        )

        with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
            workbook = writer.book
            for sheet_name, df in all_dfs.items():
                df = pd.concat(df, ignore_index=True) if df else pd.DataFrame()

                if sheet_name == "cell_data":
                    summary_df = df[["left_Node_ID", "left_MO", "left_cellId"]].copy()
                    summary_df.rename(
                        columns={
                            "left_Node_ID": "Pre SiteId",
                            "left_MO": "Pre CellName",
                            "left_cellId": "cellId",
                        },
                        inplace=True,
                    )
                    enbinfo_df = all_dfs.get("enbinfo", [pd.DataFrame()])[0]
                    enodeid_map = (
                        dict(zip(enbinfo_df["Node_ID"], enbinfo_df["eNBId"]))
                        if not enbinfo_df.empty
                        else {}
                    )
                    summary_df["Pre eNBID"] = summary_df["Pre SiteId"].map(enodeid_map)
                    summary_df.insert(2, "Pre eNBID", summary_df.pop("Pre eNBID"))

                    df.columns = [
                        col.replace("left_", "").replace("right_", "")
                        for col in df.columns
                    ]

                elif sheet_name == "gpl-para":
                    df.rename(columns={"Value": "Current value"}, inplace=True)

                elif sheet_name == "FeatureState":
                    df["featureState"] = (
                        df["featureState"].astype(str).str.split().str[0]
                    )
                    df["licenseState"] = (
                        df["licenseState"].astype(str).str.split().str[0]
                    )
                    df.rename(
                        columns={
                            "MO": "CXC ID",
                            "featureState": "Current FeatureState",
                            "licenseState": "Current LicenseState",
                        },
                        inplace=True,
                    )

                df.to_excel(writer, sheet_name=sheet_name, index=False)
                format_excel_sheet(writer, sheet_name, df)

            if "summary_df" in locals() and not summary_df.empty:
                summary_df.to_excel(writer, sheet_name="Summary", index=False)
                format_excel_sheet(writer, "Summary", summary_df)

        relative_path = os.path.relpath(excel_path, settings.MEDIA_ROOT)
        download_url = request.build_absolute_uri(
            os.path.join(settings.MEDIA_URL, relative_path)
        )
        all_download_urls.append(download_url)

    return Response(
        {"message": "Data processed successfully.", "download_urls": all_download_urls},
        status=status.HTTP_200_OK,
    )





@api_view(["POST"])
def get_pre_post_audit(request):  # noqa: F811
    ########################################################## Get pre and post files from frontend ##############################################
    pre_log_files = request.FILES.getlist("pre_log_files")
    post_log_files = request.FILES.getlist("post_log_files")

    ################################################################## Define base folders ####################################################
    base_dir = os.path.join(settings.MEDIA_ROOT, "4G_5G_GPL_AUDIT")
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    session_folder = os.path.join(
        base_dir, "GPL_AUDIT_FILES", f"GPL_AUDIT_{current_time}"
    )
    os.makedirs(session_folder, exist_ok=True, mode=0o777)

    ###################################################### Prepare folders to save raw logs #######################################################
    pre_folder = os.path.join(session_folder, "pre_logs")
    post_folder = os.path.join(session_folder, "post_logs")
    os.makedirs(pre_folder, exist_ok=True)
    os.makedirs(post_folder, exist_ok=True)

    def save_files(file_list, folder):
        saved_paths = []
        for f in file_list:
            path = os.path.join(folder, f.name)
            with open(path, "wb+") as dest:
                for chunk in f.chunks():
                    dest.write(chunk)
            saved_paths.append(path)
        return saved_paths

    #################################################################### Load command list ######################################################
    command_file_path = os.path.join(
        base_dir, "command_file", "GPL_Audit_command_updated.txt"
    )
    with open(command_file_path, "r") as command_file:
        command_lines = [line.strip() for line in command_file.readlines()]
    results = get_commands(command_lines)
    
    ################################################################## Save both pre and post files ###################################################
    pre_paths = save_files(pre_log_files, pre_folder)
    post_paths = save_files(post_log_files, post_folder)

    def extract_node(file_path):
        with open(file_path, "r") as f:
            content = f.read()
        match = re.search(r"ManagedElement=([\w-]+)", content)
        return match.group(1) if match else None

    ######################################################## Extract node names from files #################################################################
    pre_nodes = {extract_node(p): p for p in pre_paths if extract_node(p)}
    post_nodes = {extract_node(p): p for p in post_paths if extract_node(p)}

    pre_all_dfs, post_all_dfs = defaultdict(list), defaultdict(list)
    pre_all_dfs["Summary"], post_all_dfs["Summary"] = [], []

    with ThreadPoolExecutor(max_workers=3) as executor:
        # Submit PRE logs
        futures = {
            executor.submit(process_files, pre_file, pre_file, results): pre_file
            for pre_file in pre_paths
        }

        for future in futures:
            file, df_dict = future.result()
            # Explicit verification that we are handling a dictionary of dataframes
            if isinstance(df_dict, dict):
                for key, df in df_dict.items():
                    if isinstance(df, pd.DataFrame) and not df.empty:
                        pre_all_dfs[key].append(df)
            elif isinstance(df_dict, pd.DataFrame):
                # Fallback if your internal extractor methods changed signatures
                pre_all_dfs["Default"].append(df_dict)

        # Submit POST logs
        futures = {
            executor.submit(process_files, post_file, post_file, results): post_file
            for post_file in post_paths
        }

        for future in futures:
            file, df_dict = future.result()
            if isinstance(df_dict, dict):
                for key, df in df_dict.items():
                    if isinstance(df, pd.DataFrame) and not df.empty:
                        post_all_dfs[key].append(df)
            elif isinstance(df_dict, pd.DataFrame):
                post_all_dfs["Default"].append(df_dict)

    all_pre_merged_df = {
        key: pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
        for key, dfs in pre_all_dfs.items()
    }
    print("feature state for pre_df:- ", all_pre_merged_df["FeatureState"].head(10))
    all_post_merged_df = {
        key: pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
        for key, dfs in post_all_dfs.items()
    }
    print("feature state for post_df:- ", all_post_merged_df["FeatureState"].head(10))
    ################################################################ Prepare post cell data ###########################################################
    post_cell_id_df = all_post_merged_df["cell_data"][
        ["left_Node_ID", "left_MO", "left_cellId"]
    ].rename(
        columns={
            "left_Node_ID": "Post SiteId",
            "left_MO": "Post CellName",
            "left_cellId": "cellId",
        }
    )
    pre_cell_id_df = all_pre_merged_df["cell_data"][
        ["left_Node_ID", "left_MO", "left_cellId"]
    ].rename(
        columns={
            "left_Node_ID": "Pre SiteId",
            "left_MO": "Pre CellName",
            "left_cellId": "cellId",
        }
    )

    ############################################################ mapping eNBId with Site ID ############################################################
    pre_enodeid_mapping = {
        row["Node_ID"]: row["eNBId"]
        for _, row in all_pre_merged_df["enbinfo"].iterrows()
    }
    post_enodeid_mapping = {
        row["Node_ID"]: row["eNBId"]
        for _, row in all_post_merged_df["enbinfo"].iterrows()
    }

    #####################################################################################################################################################

    # ------------------------------------------------------------------ Merge pre and post data ---------------------------------------------------------#
    pre_cell_id_df["cellId"] = pre_cell_id_df["cellId"].astype(str)
    post_cell_id_df["cellId"] = post_cell_id_df["cellId"].astype(str)
    all_post_merged_df["Summary"] = pd.merge(
        pre_cell_id_df, post_cell_id_df, on=["cellId"], how="outer"
    ).fillna("NA")

    all_post_merged_df["Summary"] = all_post_merged_df["Summary"].sort_values(
        by=["Pre SiteId", "Pre CellName", "cellId"], ascending=True
    )

    ##################################################################################### Adding enode id ###############################################

    all_post_merged_df["Summary"]["Post eNBID"] = all_post_merged_df["Summary"][
        "Post SiteId"
    ].apply(lambda x: post_enodeid_mapping[x] if x != "NA" else "NA")

    all_post_merged_df["Summary"]["Pre eNBID"] = all_post_merged_df["Summary"][
        "Pre SiteId"
    ].apply(lambda x: pre_enodeid_mapping[x] if x != "NA" else "NA")

    all_post_merged_df["Summary"].insert(
        5, "Post eNBID", all_post_merged_df["Summary"].pop("Post eNBID")
    )
    all_post_merged_df["Summary"].insert(
        2, "Pre eNBID", all_post_merged_df["Summary"].pop("Pre eNBID")
    )

    gpl_pre_post_file_path = os.path.join(
        session_folder, f"GPL_AUDIT_Parameter_AReport_{current_time}.xlsx"
    )
    os.makedirs(os.path.dirname(gpl_pre_post_file_path), exist_ok=True, mode=0o777)

    async def async_write_excel(
        gpl_pre_post_file_path,
        all_post_merged_df,
        all_pre_merged_df,
        format_excel_sheet,
    ):
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(
            None,
            lambda: _write_excel_sync(
                gpl_pre_post_file_path,
                all_post_merged_df,
                all_pre_merged_df,
                format_excel_sheet,
            ),
        )

    node_name_site = (
        all_post_merged_df["cell_data"]["left_MO"]
        .apply(
            lambda x: (
                str(x).split(",")[-1].split("=")[-1].split("_")[4][:-1]
                if isinstance(x, str) and "=" in x
                else None
            )
        )
        .unique()[0]
    )

    def _write_excel_sync(
        gpl_pre_post_file_path,
        all_post_merged_df,
        all_pre_merged_df,
        format_excel_sheet,
    ):
        with pd.ExcelWriter(gpl_pre_post_file_path, engine="xlsxwriter") as writer:
            workbook = writer.book
            for sheet_name, df in all_post_merged_df.items():
                if sheet_name == "Summary":
                    df.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        startcol=9,
                        startrow=3,
                        index=False,
                    )
                    format_excel_sheet(writer, sheet_name, df, startrow=3, startcol=9)

                elif sheet_name == "gpl-para":

                    def create_new_cell_mapping(cell_data_df: pd.DataFrame):
                        return {
                            row["MO"]: row["cellId"]
                            for _, row in cell_data_df.iterrows()
                        }

                    def create_cell_mappings(summary_df):
                        return (
                            {
                                row["Pre CellName"]: row["cellId"]
                                for _, row in summary_df.iterrows()
                            },
                            {
                                row["Post CellName"]: row["cellId"]
                                for _, row in summary_df.iterrows()
                            },
                            {
                                row["Pre SiteId"]: row["Post SiteId"]
                                for _, row in summary_df.iterrows()
                                if row["Post SiteId"] != "NA"
                            },
                        )

                    def add_cell_ids(df: pd.DataFrame, col_name, mapping):
                        df.insert(
                            1,
                            "cellId",
                            df["MO"].apply(lambda x: mapping.get(x.split(",")[0], 0)),
                        )
                        return df

                    cell_id_df = all_post_merged_df["Summary"].sort_values(
                        by="Pre SiteId"
                    )
                    cell_data_df = all_post_merged_df["cell_data"]
                    pre_node_name = (
                        cell_id_df["Pre CellName"]
                        .apply(
                            lambda x: (
                                str(x).split(",")[-1].split("=")[-1].split("_")[4][:-1]
                                if isinstance(x, str) and "=" in x
                                else None
                            )
                        )
                        .unique()[0]
                    )
                    post_node_name = (
                        cell_id_df["Post CellName"]
                        .apply(
                            lambda x: (
                                str(x).split(",")[-1].split("=")[-1].split("_")[4][:-1]
                                if isinstance(x, str) and "=" in x
                                else None
                            )
                        )
                        .unique()[0]
                    )
                    gpl_pre_df = all_pre_merged_df.get(sheet_name)
                    gpl_post_df: pd.DataFrame = df.copy()
                    gpl_pre_df = gpl_pre_df.assign(
                        **{
                            "Pre-existing Value": gpl_pre_df["Value"],
                            "Current value": "",
                        }
                    )
                    pre_map, post_map, site_map = create_cell_mappings(cell_id_df)
                    post_map_cellId = create_new_cell_mapping(cell_data_df)

                    gpl_pre_df["MO"] = gpl_pre_df["MO"].apply(
                        lambda x: str(x).replace(
                            f"_{pre_node_name}", f"_{post_node_name}"
                        )
                    )

                    gpl_pre_df = add_cell_ids(gpl_pre_df, "MO", post_map_cellId)
                    gpl_post_df = add_cell_ids(gpl_post_df, "MO", post_map)

                    merged_df = pd.merge(
                        left=gpl_pre_df,
                        right=gpl_post_df,
                        how="left",
                        on=["MO", "cellId", "Parameter"],
                        indicator=True,
                    )
                    # merged_df["MO_y"] = merged_df["MO_y"].fillna(merged_df["MO_x"])
                    merged_df["Node_ID_y"] = merged_df["Node_ID_y"].fillna(
                        "Cell is not Found in Post"
                    )
                    merged_df["Current value"] = merged_df["Value_y"]
                    merged_df.drop(
                        columns=["Node_ID_x", "Value_x", "Value_y"], inplace=True
                    )
                    merged_df.rename(columns={"Node_ID_y": "Node_ID"}, inplace=True)
                    merged_df.drop_duplicates(subset=["MO", "Parameter"], inplace=True)

                    merged_df["Parameter Setting Status"] = merged_df.apply(
                        lambda row: (
                            "OK"
                            if row["Current value"] == row["Pre-existing Value"]
                            else (
                                "Missing"
                                if pd.isna(row["Current value"])
                                or pd.isna(row["Pre-existing Value"])
                                else "NOT OK"
                            )
                        ),
                        axis=1,
                    )
                    merged_df["Current value"] = merged_df["Current value"].fillna("NA")

                    def get_band(cell_name):
                        if "_F1_" in cell_name:
                            return "L2100"
                        elif "_F3_" in cell_name:
                            return "L1800"
                        elif "_F8_" in cell_name:
                            return "L900"
                        elif "_T1_" in cell_name or "_T2_" in cell_name:
                            return "L23"
                        else:
                            return "Band Error"

                    merged_df["Band"] = merged_df["MO"].apply(lambda x: get_band(x))

                    merged_df = merged_df[
                        ["Node_ID", "MO"]
                        + [
                            col
                            for col in merged_df.columns
                            if col not in ["Node_ID", "MO"]
                        ]
                    ]
                    merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    format_excel_sheet(
                        writer, sheet_name, merged_df, startrow=0, startcol=0
                    )

                elif sheet_name == "FeatureState":
                    feature_pre_df: pd.DataFrame = all_pre_merged_df.get(sheet_name)
                    feature_pre_df.rename(
                        columns={
                            "featureState": "Pre Existing FeatureState",
                            "licenseState": "Pre Existing LicenseState",
                        },
                        inplace=True,
                    )
                    cell_id_df = all_post_merged_df["Summary"]

                    feature_pre_df["Current FeatureState"] = ""

                    feature_pre_df["Pre Existing FeatureState"] = feature_pre_df[
                        "Pre Existing FeatureState"
                    ].astype(str)
                    feature_pre_df["Pre Existing LicenseState"] = feature_pre_df[
                        "Pre Existing LicenseState"
                    ].astype(str)
                    
                    # --- FIXED LAMBDAS: Safely handle floats/NaNs ---
                    feature_pre_df["Pre Existing FeatureState"] = feature_pre_df[
                        "Pre Existing FeatureState"
                    ].apply(lambda x: str(x).split(" ")[0] if pd.notna(x) and " " in str(x) else x)
                    
                    feature_pre_df["Pre Existing LicenseState"] = feature_pre_df[
                        "Pre Existing LicenseState"
                    ].apply(lambda x: str(x).split(" ")[0] if pd.notna(x) and " " in str(x) else x)

                    feature_pre_df["Current LicenseState"] = ""
                    feature_pre_df.rename(columns={"MO": "CXC ID"}, inplace=True)
                    
                    feature_post_df: pd.DataFrame = df.copy()
                    feature_post_df.rename(columns={"MO": "CXC ID"}, inplace=True)
                    
                    # --- FIXED LAMBDAS: Safely handle floats/NaNs ---
                    feature_post_df["featureState"] = feature_post_df[
                        "featureState"
                    ].apply(lambda x: str(x).split(" ")[0] if pd.notna(x) and " " in str(x) else x)
                    
                    feature_post_df["licenseState"] = feature_post_df[
                        "licenseState"
                    ].apply(lambda x: str(x).split(" ")[0] if pd.notna(x) and " " in str(x) else x)

                    pre_post_site_mapping = {
                        row["Pre SiteId"]: row["Post SiteId"]
                        for i, row in cell_id_df.iterrows()
                        if row["Post SiteId"] != "NA"
                    }

                    merged_df = pd.DataFrame()
                    for node_id in feature_pre_df["Node_ID"].unique():
                        post_node_id = pre_post_site_mapping.get(node_id, None)
                        if post_node_id is None:
                            continue

                        pre_df = feature_pre_df[feature_pre_df["Node_ID"] == node_id]
                        post_df = feature_post_df[
                            feature_post_df["Node_ID"] == post_node_id
                        ]

                        commind_df = pd.merge(
                            left=pre_df,
                            right=post_df,
                            on=["CXC ID", "description"],
                            how="left",
                        )

                        commind_df["Node_ID_y"] = commind_df["Node_ID_y"].fillna(
                            commind_df["Node_ID_x"]
                        )
                        commind_df["Node_ID_x"] = commind_df["Node_ID_y"]

                        commind_df["Current FeatureState"] = commind_df["featureState"]
                        commind_df["Current LicenseState"] = commind_df["licenseState"]

                        commind_df.rename(
                            columns={"Node_ID_x": "Node_ID"}, inplace=True
                        )
                        commind_df.drop(
                            columns=["Node_ID_y", "featureState", "licenseState"],
                            inplace=True,
                        )

                        commind_df["Current FeatureState"] = commind_df[
                            "Current FeatureState"
                        ].astype(str)
                        commind_df["Current LicenseState"] = commind_df[
                            "Current LicenseState"
                        ].astype(str)

                        merged_df = pd.concat(
                            [merged_df, commind_df], axis=0, ignore_index=True
                        )

                    def check_feature_status(row):
                        curr_feat = row["Current FeatureState"]
                        pre_feat = row["Pre Existing FeatureState"]
                        curr_lic = row["Current LicenseState"]
                        pre_lic = row["Pre Existing LicenseState"]

                        if (
                            pd.isna(curr_feat)
                            or curr_feat == "nan"
                            or pd.isna(pre_lic)
                            or pre_lic == "nan"
                        ):
                            return "Missing"
                        elif curr_feat == pre_feat and curr_lic == pre_lic:
                            return "OK"
                        else:
                            return "NOT OK"

                    merged_df["Feature setting Status"] = merged_df.apply(
                        check_feature_status, axis=1
                    )

                    merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    format_excel_sheet(
                        writer, sheet_name, merged_df, startrow=0, startcol=0
                    )

                elif sheet_name == "Eutranfrequency":
                    cell_id_df = all_post_merged_df["Summary"]

                    node_mapping = {
                        row["Pre SiteId"]: row["Post SiteId"]
                        for i, row in cell_id_df.iterrows()
                        if row["Post SiteId"] != "NA"
                    }

                    pre_freq_df = all_pre_merged_df.get(sheet_name).copy()
                    post_freq_df = df.copy()

                    merged_df = pre_freq_df.merge(
                        post_freq_df, how="left", on=["MO"]
                    ).sort_values(by=["arfcnValueEUtranDl_y"])

                    merged_df["arfcnValueEUtranDl_x"] = merged_df[
                        "arfcnValueEUtranDl_x"
                    ].astype(str)
                    merged_df["arfcnValueEUtranDl_y"] = merged_df[
                        "arfcnValueEUtranDl_y"
                    ].astype(str)
                    columns = ["arfcnValueEUtranDl", "arfcn", "arfcnValueGeranDl"]
                    for col in columns:
                        merged_df["Status"] = merged_df.apply(
                            lambda row: (
                                "OK"
                                if row[f"{col}_x"] == row[f"{col}_y"]
                                else (
                                    "Missing in Post"
                                    if pd.isna(row[f"{col}_y"])
                                    or row[f"{col}_y"] == "nan"
                                    else "NOT OK"
                                )
                            ),
                            axis=1,
                        )

                    merged_df.drop(
                        columns=[
                            "Node_ID_y",
                            "arfcnValueEUtranDl_y",
                            "arfcn_y",
                            "arfcnValueGeranDl_y",
                        ],
                        inplace=True,
                    )

                    merged_df.rename(
                        columns={
                            "Node_ID_x": "Node_ID",
                            "arfcnValueEUtranDl_x": "arfcnValueEUtranDl",
                            "arfcn_x": "arfcn",
                            "arfcnValueGeranDl_x": "arfcnValueGeranDl",
                        },
                        inplace=True,
                    )

                    merged_df["Node_ID"] = merged_df["Node_ID"].apply(
                        lambda x: (
                            node_mapping[x] if x != "NA" and x in node_mapping else "NA"
                        )
                    )

                    merged_df.sort_values(by=["MO"], inplace=True)

                    merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    format_excel_sheet(
                        writer, sheet_name, merged_df, startrow=0, startcol=0
                    )

                elif sheet_name == "EutranfreqRelation":
                    pre_freq_relation_df = all_pre_merged_df.get(sheet_name)
##################################new change - no impact of this #####################   
                    # print("\n===== PRE FREQ RELATION COLUMNS =====")
                    # print(pre_freq_relation_df.columns.tolist())

                    # print("\nQCI PARAMETERS CHECK\n")

                    # qci_cols = [
                    #     col
                    #     for col in pre_freq_relation_df.columns
                    #     if "Qci" in col or "qci" in col or "a5Thr" in col
                    # ]

                    # print("Matching columns:", qci_cols)

                    # if qci_cols:
                    #     print(pre_freq_relation_df[qci_cols].head(20))
 ##################################new change - no impact of this #####################                            
                    cell_id_df = all_post_merged_df["Summary"]
                    cell_data_df = all_post_merged_df["cell_data"]
                    post_freq_relation_df = df.copy()
                    pre_cell_mapping = {
                        row["Pre CellName"]: row["cellId"]
                        for i, row in cell_id_df.iterrows()
                    }

                    pre_post_site_mapping = {
                        row["Pre SiteId"]: row["Post SiteId"]
                        for i, row in cell_id_df.iterrows()
                        if row["Post SiteId"] != "NA" and row["Pre SiteId"] != "NA"
                    }
                    post_cell_mapping = {
                        row["Post CellName"]: row["cellId"]
                        for i, row in cell_id_df.iterrows()
                    }
                    pre_freq_relation_df.insert(1, "cellId", "")
                    pre_freq_relation_df["cellId"] = pre_freq_relation_df["MO"].apply(
                        lambda x: pre_cell_mapping[x.split(",")[0] if "," in x else x]
                    )
                    pre_freq_relation_df["lbBnrPolicy"] = pre_freq_relation_df["lbBnrPolicy"].apply(lambda x: str(x).split(" ")[0] if pd.notna(x) and " " in str(x) else x)

                    post_freq_relation_df["lbBnrPolicy"] = post_freq_relation_df["lbBnrPolicy"].apply(lambda x: str(x).split(" ")[0] if pd.notna(x) and " " in str(x) else x)
                    post_freq_relation_df.insert(1, "cellId", "")
                    post_freq_relation_df["cellId"] = post_freq_relation_df["MO"].apply(
                        lambda x: post_cell_mapping[x.split(",")[0] if "," in x else x]
                    )
                    post_freq_relation_df["lbBnrPolicy"] = post_freq_relation_df[
                        "lbBnrPolicy"
                    ].apply(lambda x: str(x).split(" ")[0] if pd.notna(x) and " " in str(x) else x)

                    int64_columns = pre_freq_relation_df.select_dtypes(
                        include="int64"
                    ).columns.tolist()
                    for col in int64_columns:
                        post_freq_relation_df[col] = post_freq_relation_df[col].astype(
                            "int64"
                        )

                    pre_node_name = (
                        cell_id_df["Pre CellName"]
                        .apply(
                            lambda x: (
                                str(x).split(",")[-1].split("=")[-1].split("_")[4][:-1]
                                if isinstance(x, str) and "=" in x
                                else None
                            )
                        )
                        .unique()[0]
                    )
                    post_node_name = (
                        cell_id_df["Post CellName"]
                        .apply(
                            lambda x: (
                                str(x).split(",")[-1].split("=")[-1].split("_")[4][:-1]
                                if isinstance(x, str) and "=" in x
                                else None
                            )
                        )
                        .unique()[0]
                    )

                    pre_freq_relation_df["MO"] = pre_freq_relation_df["MO"].apply(
                        lambda x: str(x).replace(pre_node_name, post_node_name)
                    )

                    post_cellId_mapping = {
                        row["MO"]: row["cellId"] for i, row in cell_data_df.iterrows()
                    }

                    post_cellId_node_mapping = {
                        row["cellId"]: row["Node_ID"]
                        for i, row in cell_data_df.iterrows()
                    }
                    print("post mo cell mapping with this:- \n", post_cellId_mapping)

                    pre_freq_relation_df["cellId"] = pre_freq_relation_df["MO"].apply(
                        lambda x: post_cellId_mapping.get(x.split(",")[0], "")
                    )
#####################new change -important Duplicate remover added#######################
                    pre_freq_relation_df = pre_freq_relation_df.drop_duplicates(
                        subset=["MO", "cellId"]
                    )

                    post_freq_relation_df = post_freq_relation_df.drop_duplicates(
                        subset=["MO", "cellId"]
                    )
#####################new change -important Duplicate remover added#######################
                    merged_df = pd.merge(
                        left=pre_freq_relation_df,
                        right=post_freq_relation_df,
                        on=["MO", "cellId"],
                        how="left",
                        suffixes=("_x", "_y"),
                    )
                    
#####################new change - Important Float64 def #######################
                    
                    for col in merged_df.columns:
                        merged_df[col] = merged_df[col].fillna("")

                    float64_to_int64 = merged_df.select_dtypes(
                        include="float64"
                    ).columns.tolist()
                    float64_to_int64 = merged_df.select_dtypes(
                        include="float64"
                    ).columns.tolist()
                    
#####################new change - Important Float64 def #######################
                                        
                    for column in float64_to_int64:
                        merged_df[column] = (
                            pd.to_numeric(merged_df[column], errors="coerce")
                            .replace([float("inf"), float("-inf")], pd.NA)
                            .astype("Int64")
                        )
                    merged_df["eutranFrequencyRef_y"] = merged_df[
                        "eutranFrequencyRef_y"
                    ].fillna(merged_df["eutranFrequencyRef_x"])
                    merged_df.drop(columns=["eutranFrequencyRef_x"], inplace=True)
                    merged_df.rename(
                        columns={"eutranFrequencyRef_y": "eutranFrequencyRef"},
                        inplace=True,
                    )
                    merged_df["Node_ID_y"] = merged_df["Node_ID_y"].fillna(
                        "Cell is not Found in Post"
                    )
                    merged_df["Node_ID_x"] = merged_df["Node_ID_y"]
                    merged_df["duplicates_mask"] = merged_df.duplicated(
                        subset=merged_df.columns.tolist()
                    )
                    merged_df.drop(columns=["Node_ID_y"], inplace=True)
                    merged_df.rename(columns={"Node_ID_x": "Node_ID"}, inplace=True)

                    merged_df.insert(3, "Status", "OK")
                    for col in pre_freq_relation_df.columns:
                        if col not in [
                            "MO",
                            "Node_ID",
                            "eutranFrequencyRef",
                            "cellId",
                        ]:
                            pre_col = f"{col}_x"
                            post_col = f"{col}_y"
                                            
#####################new change - important Fill NaN values with empty strings#######################
                            ##old##
                            # merged_df[pre_col] = (
                            #     merged_df[pre_col].astype(str).str.lower()
                            # )
                            # merged_df[post_col] = (
                            #     merged_df[post_col].astype(str).str.lower()
                            # )     
                            #old###
                            
                            ## new ###                    
                            merged_df[pre_col] = (
                                merged_df[pre_col]
                                .fillna("")
                                .astype(str)
                                .replace("nan", "")
                                .str.lower()
                            )
                            merged_df[post_col] = (
                                merged_df[post_col]
                                .fillna("")
                                .astype(str)
                                .replace("nan", "")
                                .str.lower()
                            )
                            ###new ###
#####################new change - important Fill NaN values with empty strings#######################
                            
                            mask = merged_df[pre_col].ne(merged_df[post_col])
                            merged_df.loc[mask, "Status"] = "NOT OK"
                            for idx in merged_df[mask].index:
                                pre_value = merged_df.at[idx, pre_col]
                                post_value = merged_df.at[idx, post_col]
                                if pd.isna(post_value) or post_value in [
                                    "nan",
                                    0,
                                    "0",
                                    "<na>",
                                    "<NA>",
                                ]:
                                    merged_df.at[idx, pre_col] = f"{pre_value}"
                                else:
                                    merged_df.at[idx, pre_col] = (
                                        f"{pre_value}|{post_value}"
                                    )

                    merged_df["Status"] = merged_df.apply(
                        lambda row: (
                            "Missing"
                            if row["Node_ID"] == "Cell is not Found in Post"
                            else row["Status"]
                        ),
                        axis=1,
                    )
                    merged_df["Node_ID"] = merged_df["cellId"].apply(
                        lambda x: post_cellId_node_mapping.get(x.split(",")[0], "")
                    )
                    merged_df.sort_values(by=["Node_ID"], inplace=True)

                    merged_df = merged_df[
                        [
                            col if col.endswith("_x") else col
                            for col in merged_df.columns
                            if not col.endswith("_y")
                        ]
                    ]
                    merged_df.rename(
                        columns={
                            col: col.replace("_x", "") for col in merged_df.columns
                        },
                        inplace=True,
                    )
                    merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    format_excel_sheet(
                        writer, sheet_name, merged_df, startrow=0, startcol=0
                    )

                elif sheet_name == "CellRelation":
                    # ---------------------------------------------------------------------------------------------------------------------------------------#
                    pre_cell_relation_df = all_pre_merged_df.get(sheet_name)
                    post_cell_relaton_df = all_post_merged_df.get(sheet_name)
                    pre_cell_id_mapping_with_cell = {
                        row["Pre CellName"]: row["cellId"]
                        for _, row in pre_cell_id_df.iterrows()
                    }
                    reversed_pre_cell_id_mapping_with_cell = {
                        v: k for k, v in pre_cell_id_mapping_with_cell.items()
                    }
                    post_cell_id_mapping_with_cell = {
                        row["Post CellName"]: row["cellId"]
                        for _, row in post_cell_id_df.iterrows()
                    }
                    reversed_post_cell_id_mapping_with_cell = {
                        v: k for k, v in post_cell_id_mapping_with_cell.items()
                    }
                    cell_id_mapped_with_nodeId = {
                        row["cellId"]: row["Post SiteId"]
                        for _, row in post_cell_id_df.iterrows()
                    }
                    # ------------------------------------------------------------ mapping the siteId with the cellId --------------------------------------------------#
                    post_site_mapped_with_cellId = {
                        row["cellId"]: row["Post SiteId"]
                        for _, row in post_cell_id_df.iterrows()
                    }
                    # -------------------------########################## pre cell relation with cellId #########################################-------------#
                    if "noughbourCellId" not in pre_cell_relation_df.columns:
                        pre_cell_relation_df.insert(2, "noughbourCellId", "")
                    if "cellId" not in pre_cell_relation_df.columns:
                        pre_cell_relation_df.insert(3, "cellId", "")

                    if "noughbourCellId" not in post_cell_relaton_df.columns:
                        post_cell_relaton_df.insert(2, "noughbourCellId", "")
                    if "cellId" not in post_cell_relaton_df.columns:
                        post_cell_relaton_df.insert(3, "cellId", "")

                    pre_cell_relation_df["noughbourCellId"] = pre_cell_relation_df[
                        "MO"
                    ].apply(lambda mo: str(mo).split("-")[-1])

                    pre_cell_relation_df["cellId"] = pre_cell_relation_df["MO"].apply(
                        lambda mo: pre_cell_id_mapping_with_cell.get(
                            str(mo).split(",")[0]
                        )
                    )
                    post_cell_relaton_df["noughbourCellId"] = post_cell_relaton_df[
                        "MO"
                    ].apply(lambda mo: str(mo).split("-")[-1])
                    post_cell_relaton_df["cellId"] = post_cell_relaton_df["MO"].apply(
                        lambda mo: post_cell_id_mapping_with_cell.get(
                            str(mo).split(",")[0]
                        )
                    )

                    def extract_enb_id(mo: str) -> str:
                        try:
                            return mo.split("-")[1]
                        except (AttributeError, IndexError):
                            return None

                    pre_eNBID = (
                        all_pre_merged_df.get("enbinfo")["eNBId"].unique().tolist()
                    )
                    post_eNBID = (
                        all_post_merged_df["enbinfo"]["eNBId"].unique().tolist()
                    )

                    pre_cell_relation_df["eNBId"] = pre_cell_relation_df["MO"].apply(
                        extract_enb_id
                    )
                    post_cell_relaton_df["eNBId"] = post_cell_relaton_df["MO"].apply(
                        extract_enb_id
                    )

                    pre_cell_relation_df = pre_cell_relation_df[
                        pre_cell_relation_df["eNBId"].isin(pre_eNBID)
                    ]
                    post_cell_relaton_df = post_cell_relaton_df[
                        post_cell_relaton_df["eNBId"].isin(post_eNBID)
                    ]

                    pre_cell_relation_df.drop(columns=["eNBId"], inplace=True)
                    post_cell_relaton_df.drop(columns=["eNBId"], inplace=True)
                    # --------------------------------------------------------- find pre and post node identification----------------------------------#
                    pre_node_name = (
                        pre_cell_id_df["Pre CellName"]
                        .apply(
                            lambda x: (
                                str(x).split(",")[-1].split("=")[-1].split("_")[4][:-1]
                                if isinstance(x, str) and "=" in x
                                else None
                            )
                        )
                        .unique()[0]
                    )
                    post_node_name = (
                        post_cell_id_df["Post CellName"]
                        .apply(
                            lambda x: (
                                str(x).split(",")[-1].split("=")[-1].split("_")[4][:-1]
                                if isinstance(x, str) and "=" in x
                                else None
                            )
                        )
                        .unique()[0]
                    )

                    columns_to_convert = [
                        "cellIndividualOffsetEUtran",
                        "coverageIndicator",
                        "loadBalancing",
                        "qOffsetCellEUtran",
                        "reportDlActivity",
                        "sCellCandidate",
                        "sCellPriority",
                        "sleepModeCovCellCandidate",
                    ]

                    def convert_to_int(value):
                        try:
                            if pd.isna(value):
                                return None
                            value = (
                                str(value).split(" ")[0] if " " in str(value) else value
                            )
                            return int(value)
                        except Exception:
                            return None

                    for columns in columns_to_convert:
                        pre_cell_relation_df[columns] = pre_cell_relation_df[
                            columns
                        ].apply(lambda x: str(x).split(" ")[0] if " " in str(x) else x)
                        post_cell_relaton_df[columns] = post_cell_relaton_df[
                            columns
                        ].apply(lambda x: str(x).split(" ")[0] if " " in str(x) else x)
                        pre_cell_relation_df[columns] = pre_cell_relation_df[
                            columns
                        ].apply(convert_to_int)
                        post_cell_relaton_df[columns] = post_cell_relaton_df[
                            columns
                        ].apply(convert_to_int)

                    pre_cell_relation_df["cellId"] = pre_cell_relation_df[
                        "cellId"
                    ].apply(lambda x: reversed_pre_cell_id_mapping_with_cell.get(x, ""))

                    pre_cell_relation_df["noughbourCellId"] = pre_cell_relation_df[
                        "noughbourCellId"
                    ].apply(lambda x: reversed_pre_cell_id_mapping_with_cell.get(x, ""))

                    pre_cell_relation_df["cellId"] = pre_cell_relation_df[
                        "cellId"
                    ].apply(lambda x: str(x).replace(pre_node_name, post_node_name))
                    pre_cell_relation_df["MO"] = pre_cell_relation_df["MO"].apply(
                        lambda x: str(x).replace(pre_node_name, post_node_name)
                    )
                    pre_cell_relation_df["cellId"] = pre_cell_relation_df[
                        "cellId"
                    ].apply(lambda x: post_cell_id_mapping_with_cell.get(x))
                    pre_cell_relation_df["noughbourCellId"] = pre_cell_relation_df[
                        "noughbourCellId"
                    ].apply(lambda x: str(x).replace(pre_node_name, post_node_name))
                    pre_cell_relation_df["noughbourCellId"] = pre_cell_relation_df[
                        "noughbourCellId"
                    ].apply(lambda x: post_cell_id_mapping_with_cell.get(x))
                    print(post_cell_id_mapping_with_cell)

                    pre_cell_relation_df["Node_ID"] = pre_cell_relation_df[
                        "cellId"
                    ].apply(lambda x: post_site_mapped_with_cellId.get(x, ""))

                    # ------------------------------------------------------- changing the MO columns nouhbour cel --------------------------------------------------------#
                    def replace_mo_nighbourCell(row):
                        mo_str = row["MO"]
                        noughbour_cell_id = row["noughbourCellId"]
                        cell_id_from_mo = mo_str.split("-")[-1]

                        return (
                            mo_str.replace(cell_id_from_mo, noughbour_cell_id)
                            if noughbour_cell_id
                            else mo_str
                        )

                    pre_cell_relation_df["MO"] = pre_cell_relation_df.apply(
                        lambda row: replace_mo_nighbourCell(row), axis=1
                    )

                    def replace_all_enbid(mo):
                        mo_str = str(mo)
                        for pre, post in zip(pre_eNBID, post_eNBID):
                            mo_str = mo_str.replace(f"-{pre}-", f"-{post}-")
                        return mo_str

                    pre_cell_relation_df["MO"] = pre_cell_relation_df["MO"].apply(
                        replace_all_enbid
                    )
                    # ------------------------------------------------------------------------------------------------------------------------------------------------------#
                    merged_df = pd.merge(
                        left=pre_cell_relation_df,
                        right=post_cell_relaton_df,
                        how="left",
                        on=["MO", "cellId"],
                    ).drop_duplicates(subset=["MO", "cellId"])
                    merged_df["neighborCellRef_y"] = merged_df[
                        "neighborCellRef_y"
                    ].fillna(merged_df["neighborCellRef_x"])
                    merged_df["neighborCellRef_x"] = merged_df["neighborCellRef_y"]
                    merged_df["noughbourCellId_y"] = merged_df[
                        "noughbourCellId_y"
                    ].fillna(merged_df["noughbourCellId_x"])
                    merged_df["noughbourCellId_x"] = merged_df["noughbourCellId_y"]

                    merged_df["Node_ID_y"] = merged_df["Node_ID_y"].fillna(
                        "cell not found in post"
                    )
                    merged_df["Node_ID_x"] = merged_df["Node_ID_y"]

                    merged_df.rename(
                        columns={
                            "Node_ID_x": "Node_ID",
                            "neighborCellRef_x": "neighborCellRef",
                            "noughbourCellId_x": "noughbourCellId",
                        },
                        inplace=True,
                    )

                    for col in columns_to_convert:
                        merged_df[f"{col}_x"] = pd.to_numeric(
                            merged_df[f"{col}_x"], errors="coerce"
                        )
                        merged_df[f"{col}_x"] = (
                            merged_df[f"{col}_x"].fillna(0).astype(int)
                        )
                        merged_df[f"{col}_y"] = pd.to_numeric(
                            merged_df[f"{col}_y"], errors="coerce"
                        )
                        merged_df[f"{col}_y"] = (
                            merged_df[f"{col}_y"].fillna(0).astype(int)
                        )

                    merged_df.insert(3, "Status", "OK")

                    for col in pre_cell_relation_df.columns:
                        if col not in [
                            "MO",
                            "Node_ID",
                            "neighborCellRef",
                            "cellId",
                            "noughbourCellId",
                        ]:
                            pre_col = f"{col}_x"
                            post_col = f"{col}_y"

                            merged_df[pre_col] = (
                                merged_df[pre_col].astype(str).str.lower()
                            )
                            merged_df[post_col] = (
                                merged_df[post_col].astype(str).str.lower()
                            )

                            mask = merged_df[pre_col].ne(merged_df[post_col])

                            merged_df.loc[mask, "Status"] = "NOT OK"
                            for idx in merged_df[mask].index:
                                pre_value = merged_df.at[idx, pre_col]
                                post_value = merged_df.at[idx, post_col]

                                if pd.isna(post_value) or post_value in ["nan"]:
                                    merged_df.at[idx, pre_col] = f"{pre_value}"

                                else:
                                    merged_df.at[idx, pre_col] = (
                                        f"{pre_value}|{post_value}"
                                    )

                    merged_df["Status"] = merged_df.apply(
                        lambda row: (
                            "Missing"
                            if row["Node_ID"] == "cell not found in post"
                            else row["Status"]
                        ),
                        axis=1,
                    )

                    merged_df["Node_ID"] = merged_df["cellId"].apply(
                        lambda x: cell_id_mapped_with_nodeId.get(x, "")
                    )
                    merged_df.sort_values(by=["Node_ID"], inplace=True)
                    merged_df = merged_df[
                        [
                            col if col.endswith("_x") else col
                            for col in merged_df.columns
                            if not col.endswith("_y")
                        ]
                    ]
                    #
                    merged_df.rename(
                        columns={
                            col: col.replace("_x", "") for col in merged_df.columns
                        },
                        inplace=True,
                    )
#####################new change - Important extra column remover #######################
                    
                    columns_to_remove = [
                        "earfcndl",
                        "physicalLayerCellIdGroup",
                        "physicalLayerSubCellId",
                        "tac",
                    ]

                    merged_df.drop(
                        columns=[c for c in columns_to_remove if c in merged_df.columns],
                        inplace=True,
                    )
 #####################new change - Important extra column remover #######################
                   
                    merged_df["neighborCellRef"] = merged_df["noughbourCellId"].apply(
                        lambda x: reversed_post_cell_id_mapping_with_cell.get(x, "")
                    )
                    for col in columns_to_convert:
                        merged_df.loc[merged_df["Status"] == "Missing", col] = (
                            merged_df.loc[merged_df["Status"] == "Missing", col].apply(
                                lambda x: (
                                    int(str(x).split("|")[0])
                                    if "|" in str(x)
                                    else (int(x) if str(x).isdigit() else x)
                                )
                            )
                        )

                    # merged_df.sort_values(by=["MO", "Node_ID"], inplace=True)
                    #
                    merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    format_excel_sheet(
                        writer, sheet_name, merged_df, startrow=0, startcol=0
                    )
                
                else:
                    if sheet_name == "cell_data":
                        df.rename(
                            columns={
                                "right_MO": "right_SectorCarrierID",
                                "right_sectorFunctionRef": "right_SectorId",
                            },
                            inplace=True,
                        )

                        df.columns = [
                            col.replace("left_", "")
                            if col.startswith("left_")
                            else col.replace("right_", "")
                            for col in df.columns
                        ]

                        required_columns = [
                            "Node_ID",
                            "MO",
                            "administrativeState",
                            "cellId",
                            "cellSubscriptionCapacity",
                            "channelBandwidth",
                            "dlChannelBandwidth",
                            "earfcn",
                            "earfcndl",
                            "earfcnul",
                            "freqBand",
                            "noOfPucchSrUsers",
                            "operationalState",
                            "physicalLayerCellId",
                            "physicalLayerCellIdGroup",
                            "physicalLayerSubCellId",
                            "primaryPlmnReserved",
                            "rachRootSequence",
                            "tac",
                            "ulChannelBandwidth",
                            "userLabel",
                            "SectorId",
                            "SectorCarrierID",
                            "configuredMaxTxPower",
                            "noOfRxAntennas",
                            "noOfTxAntennas",
                        ]

                        df = df[[col for col in required_columns if col in df.columns]]

                        columns_to_convert = [
                            "dlChannelBandwidth",
                            "channelBandwidth",
                            "earfcndl",
                            "earfcn",
                        ]
                        for col in columns_to_convert:
                            if col in df.columns:
                                df[col] = df[col].fillna("").astype(str)

                        if (
                            "dlChannelBandwidth" in df.columns
                            and "channelBandwidth" in df.columns
                        ):
                            df["dlChannelBandwidth"] = (
                                df["dlChannelBandwidth"] + df["channelBandwidth"]
                            )
                            df.drop(columns=["channelBandwidth"], inplace=True)

                        if "earfcndl" in df.columns and "earfcn" in df.columns:
                            df["earfcndl"] = df["earfcndl"] + df["earfcn"]
                            df.drop(columns=["earfcn"], inplace=True)

                        df.replace(["nan", pd.NA], "", inplace=True)

                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    format_excel_sheet(writer, sheet_name, df, startrow=0, startcol=0)

    ################################################################################## After writing the Excel file ##########################################################################
    asyncio.run(
        async_write_excel(
            gpl_pre_post_file_path,
            all_post_merged_df,
            all_pre_merged_df,
            format_excel_sheet,
        )
    )

    ################################################################################### Now it's safe to read the file ########################################################################

    # -------------------------------------------------------------------- GPL Correction File Script Generation ------------------------------------------------------------------#

    async def async_generate_correction_scripts(
        session_folder, current_time, post_nodes, gpl_pre_post_file_path
    ):
        gpl_correction_data_file = pd.ExcelFile(gpl_pre_post_file_path)
        gpl_correction_data_file_df = await asyncio.to_thread(
            gpl_correction_data_file.parse, "gpl-para"
        )
        feature_correctin_data_file_df = await asyncio.to_thread(
            gpl_correction_data_file.parse, "FeatureState"
        )
        feature_correctin_data_f_df = feature_correctin_data_file_df[
            feature_correctin_data_file_df["Feature setting Status"] == "NOT OK"
        ].copy()
        eutranFreq_correctin_data_file = await asyncio.to_thread(
            gpl_correction_data_file.parse, "Eutranfrequency"
        )
        eutranFreq_correctin_data_file_df = eutranFreq_correctin_data_file[
            eutranFreq_correctin_data_file["Status"] == "Missing in Post"
        ].copy()
        eutranFreqRelation_correctin_data_file = await asyncio.to_thread(
            gpl_correction_data_file.parse, "EutranfreqRelation"
        )
        eutranFreqRelation_correctin_data_f_df = eutranFreqRelation_correctin_data_file[
            eutranFreqRelation_correctin_data_file["Status"] != "OK"
        ].copy()
        eutranFreqRelation_cellRelation_file = await asyncio.to_thread(
            gpl_correction_data_file.parse, "CellRelation"
        )
        eutranFreqRelation_cellRelation_file = eutranFreqRelation_cellRelation_file[
            eutranFreqRelation_cellRelation_file["Status"] != "OK"
        ].copy()
        mask = (
            gpl_correction_data_file_df["Parameter Setting Status"] != "OK"
        ) & pd.notna(gpl_correction_data_file_df["Current value"])
        correctin_data_file_df = gpl_correction_data_file_df[mask].copy()

        cell_data_df = await asyncio.to_thread(
            gpl_correction_data_file.parse, "cell_data"
        )
        # Safely check for earfcndl, fallback to earfcn, or default to an empty list
        if "earfcndl" in cell_data_df.columns:
            valid_frequencies = cell_data_df["earfcndl"].dropna().unique().tolist()
        elif "earfcn" in cell_data_df.columns:
            valid_frequencies = cell_data_df["earfcn"].dropna().unique().tolist()
        else:
            valid_frequencies = []
        tasks = []
        for node_name in post_nodes:
            gpl_correction_file_path = os.path.join(
                session_folder, f"{node_name}_GPL_Correction_Script_{current_time}.txt"
            )
            gpl_correctin_df = correctin_data_file_df[
                correctin_data_file_df["Node_ID"] == node_name
            ].copy()
            feature_correctin_df = feature_correctin_data_f_df[
                feature_correctin_data_f_df["Node_ID"] == node_name
            ].copy()
            eutranFreqRel_correctin_df = eutranFreq_correctin_data_file_df[
                eutranFreq_correctin_data_file_df["Node_ID"] == node_name
            ].copy()
            eutranFreqRelation_correctin_df = eutranFreqRelation_correctin_data_f_df[
                eutranFreqRelation_correctin_data_f_df["Node_ID"].isin([node_name])
            ].copy()
            eutranFreqRelationCellRelation_df = eutranFreqRelation_cellRelation_file[
                eutranFreqRelation_cellRelation_file["Node_ID"].isin(
                    [node_name, "Cell is not Found in Post"]
                )
            ].copy()

            print("eutrancell_relation:- \n", eutranFreqRelationCellRelation_df)
            gpl_commands = []
            gpl_commands.append(
                f"##################### GPL Parameter Correction Commands {node_name}####################"
            )
            for _, row in gpl_correctin_df.iterrows():
                if pd.notna(row.get("Pre-existing Value")) and pd.notna(row.get("MO")):
                    value = str(row["Pre-existing Value"]).split()[0]
                    gpl_commands.append(f"set {row['MO']}$ {row['Parameter']} {value}")

            gpl_commands.append(
                f"##################### LTE Feature Correction Commands {node_name}####################"
            )
            for _, row in feature_correctin_df.iterrows():
                if pd.notna(row.get("Pre Existing FeatureState")) and pd.notna(
                    row.get("CXC ID")
                ):
                    value = int(row["Pre Existing FeatureState"])
                    gpl_commands.append(f"set {row['CXC ID']}$ featureState {value}")
            # Use asyncio.to_thread to run blocking IO in thread pool
            tasks.append(
                asyncio.to_thread(
                    process_correction_script_generation,
                    gpl_commands,
                    output_file_path=os.path.join(
                        session_folder, gpl_correction_file_path
                    ),
                )
            )

            eutranFreqRel_correctin_df_path = os.path.join(
                session_folder,
                f"{node_name}_Relation_Correction_Script_{current_time}.txt",
            )

            # --------------------------------------------------------------------------------------- cell relation mapping with freq ----------------------------------------------------------------#
            cell_id_df = gpl_correction_data_file.parse("cell_data")

            # Safely determine which frequency column exists in the dataframe
            if "earfcndl" in cell_id_df.columns:
                freq_col = "earfcndl"
            elif "earfcn" in cell_id_df.columns:
                freq_col = "earfcn"
            else:
                freq_col = None

            # Perform the grouping only if the column actually exists
            if freq_col:
                freq_rel_dict = cell_id_df.groupby(freq_col)["MO"].apply(list).to_dict()
            else:
                freq_rel_dict = {}  # Fallback to an empty dict if no frequency data is present
            freq_mo_names = []
            for key, value in freq_rel_dict.items():
                freq_mo_names.extend(value)

            print(freq_mo_names)
            # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
            eutran_freq_relation_script = ""
            print(eutranFreqRel_correctin_df)

            for _, row in eutranFreqRel_correctin_df.iterrows():
                mo = str(row.get("MO", ""))
                if mo.startswith("GeraNetwork=1,GeranFrequency="):
                    arfcn_val = int(row.get("arfcnValueGeranDl", ""))

                    eutran_freq_relation_script += (
                        GeranFrequency_defination.format(arfcnValueGeranDl=arfcn_val)
                        + "\n"
                    )
                elif mo.startswith("ENodeBFunction=1,EUtraNetwork=1,EUtranFrequency="):
                    arfcn_val = int(row.get("arfcnValueEUtranDl", ""))
                    eutran_freq_relation_script += (
                        EUtranFrequency_Definition.format(arfcnValueEUtranDl=arfcn_val)
                        + "\n"
                    )
            # tasks.append(asyncio.to_thread(create_freqency_relation_script, eutran_freq_relation_script, output_file_path=eutranFreqRel_correctin_df_path))
            # --------------------------------------------------------------------- frequency relation script generation ------------------------------------------------------------------#
            not_ok_freq_relation_df = eutranFreqRelation_correctin_df[
                eutranFreqRelation_correctin_df["Status"] == "NOT OK"
            ].copy()
            missing_freq_df = eutranFreqRelation_correctin_df[
                eutranFreqRelation_correctin_df["Status"] == "Missing"
            ].copy()
            print("first time:- \n", missing_freq_df)
            not_ok_freq_relation_df.drop(
                columns=["cellId", "Status", "Node_ID", "duplicates_mask"], inplace=True
            )

            missing_freq_df.drop(
                columns=["cellId", "Status", "Node_ID", "duplicates_mask"], inplace=True
            )

            id_vars = ["MO"]
            value_vars = [
                col for col in not_ok_freq_relation_df.columns if col not in id_vars
            ]
            not_ok_freq_relation_df = not_ok_freq_relation_df.melt(
                id_vars=id_vars,
                value_vars=value_vars,
                var_name="Relation Parameter",
                value_name="Value",
            )
            not_ok_freq_relation_df.sort_values(by="MO", inplace=True)
            not_ok_freq_relation_df = not_ok_freq_relation_df[
                (
                    not_ok_freq_relation_df["Value"]
                    .astype(str)
                    .str.contains(r"\|", regex=True, na=False)
                )
            ]
            not_ok_freq_relation_df["Value"] = not_ok_freq_relation_df["Value"].apply(
                lambda x: str(x).split("|")[0] if "|" in str(x) else x
            )
            print(not_ok_freq_relation_df)
            id_vars = ["MO"]
            value_vars = [
                col
                for col in missing_freq_df.columns
                if col not in id_vars and col not in ["eutranFrequencyRef", "Node_ID"]
            ]
            print(missing_freq_df)
            missing_freq_df = missing_freq_df.melt(
                id_vars=id_vars,
                value_vars=value_vars,
                var_name="Relation Parameter",
                value_name="Value",
            )

            missing_freq_df = missing_freq_df[~pd.isna(missing_freq_df["MO"])]

            missing_freq_df.drop_duplicates(
                subset=["MO", "Relation Parameter"], inplace=True
            )

            setting_eutranRelation = []
            setting_eutranRelation.append(
                f"\n\n##################### Eutran Frequency Relation Correction Commands {node_name} ####################\n\n"
            )
            for _, row in not_ok_freq_relation_df.iterrows():
                if pd.notna(row.get("MO")):
                    mo = str(row["MO"])
                    relation_parameter = row["Relation Parameter"]
                    value = str(row["Value"])
                    setting_eutranRelation.append(
                        f"set {mo}$ {relation_parameter} {value}"
                    )
            setting_eutranRelation.append(
                f"\n\n##################### Creating Eutran Frequency Relation Missing Commands post node: {node_name} ####################\n\n"
            )

            def generate_crn_section(df: pd.DataFrame):
                mos_content = []
                if not df.empty:
                    df = df.dropna(subset=["MO", "Relation Parameter", "Value"])
                    for mo, group in df.groupby("MO"):
                        freq = mo.split(",")[1].split("=")[1]
                        if freq in valid_frequencies:
                            mos_content.append(f"\ncrn {mo}")
                            for _, row in group.iterrows():
                                mos_content.append(
                                    f"{row['Relation Parameter']} {row['Value']}"
                                )
                            mos_content.append("end")
                return mos_content

            mos_content = generate_crn_section(missing_freq_df)
            setting_eutranRelation.extend(mos_content)
            setting_eutranRelation.append(
                f"\n\n\n\n######################################################## Cell Relation {node_name} ####################################################\n\n\n\n"
            )
            # ---------------------------------------------------------------------------------------------- CELL RELATION -------------------------------------------------------------------------------#
            not_ok_cell_relation = eutranFreqRelationCellRelation_df[
                eutranFreqRelationCellRelation_df["Status"] == "NOT OK"
            ].copy()
            not_ok_cell_relation = not_ok_cell_relation[
                [
                    "MO",
                    "Node_ID",
                    "Status",
                    "cellIndividualOffsetEUtran",
                    "coverageIndicator",
                    "loadBalancing",
                    "qOffsetCellEUtran",
                    "reportDlActivity",
                    "sCellCandidate",
                    "sCellPriority",
                    "sleepModeCovCellCandidate",
                ]
            ]
            # ***********************************************
            missing_cell_relation_df = eutranFreqRelationCellRelation_df[
                eutranFreqRelationCellRelation_df["Status"] == "Missing"
            ].copy()
            not_ok_cell_relation.drop(columns=["Status", "Node_ID"], inplace=True)

            missing_cell_relation_df.drop(columns=["Status", "Node_ID"], inplace=True)
            id_vars = ["MO"]
            value_vars = [
                col for col in not_ok_cell_relation.columns if col not in id_vars
            ]
            not_ok_cell_relation = not_ok_cell_relation.melt(
                id_vars=id_vars,
                value_vars=value_vars,
                var_name="Relation Parameter",
                value_name="Value",
            )

            not_ok_cell_relation["Value"] = not_ok_cell_relation["Value"].apply(
                lambda x: str(x).split("|")[0] if "|" in str(x) else x
            )

            not_ok_cell_relation.drop_duplicates(
                subset=["MO", "Relation Parameter"], inplace=True
            )
            # --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
            print("mo relation df\n", not_ok_cell_relation)
 ##############################New Change - no impact ##################################             
            # print("\nALL RELATION PARAMETERS\n")
            # print(
            #     not_ok_cell_relation["Relation Parameter"]
            #     .drop_duplicates()
            #     .sort_values()
            #     .tolist()
            # )
##############################New Change - no impact  ##################################            
            for _, row in not_ok_cell_relation.iterrows():
                mo_cell = row.get("MO", "")
                relation_parameter = row.get("Relation Parameter", "")
                value = int(row.get("Value", 0))
                setting_eutranRelation.append(
                    f"set {mo_cell}$ {relation_parameter} {value}"
                )
            # ------------------------------------------------------------------------------------- cell relation creation -------------------------------------------------------------------------------#
            setting_eutranRelation.append(
                f"\n\n\n\n###############################################CELL RELATION {node_name} #########################################################\n\n\n\n"
            )
            for _, row in missing_cell_relation_df.iterrows():
                mo_cell = row.get("MO", "")
                nighbour_cell = row.get("neighborCellRef", "")
                setting_eutranRelation.append(
                    Eutran_freq_cell_relation_defination.format(
                        EUtranCellName=mo_cell, neighborCellRef=nighbour_cell
                    )
                )

            tasks.append(
                asyncio.to_thread(
                    create_freqency_relation_script,
                    eutran_freq_relation_script,
                    setting_eutranRelation,
                    output_file_path=eutranFreqRel_correctin_df_path,
                )
            )
        await asyncio.gather(*tasks)

    #######----------------------------------------########################### Usage in your view (example, must be called from an async context): #########------------------------------------------------------#############################
    asyncio.run(
        async_generate_correction_scripts(
            session_folder, current_time, post_nodes, gpl_pre_post_file_path
        )
    )
    # ------------------------------------------------------------------------- Generate download URL -----------------------------------------------------------------------------#

    ####################################################### Create a zip file of the session_folder (which contains the Excel file and scripts)
    zip_filename = f"{node_name_site}_GPL_AUDIT_{current_time}.zip"
    zip_filepath = os.path.join(session_folder, zip_filename)
    with zipfile.ZipFile(zip_filepath, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(session_folder):
            for file in files:
                if file != zip_filename:
                    abs_path = os.path.join(root, file)
                    rel_path = os.path.relpath(abs_path, session_folder)
                    zipf.write(abs_path, rel_path)

    # Generate download URL for the zip file
    relative_url = zip_filepath.replace(str(settings.MEDIA_ROOT), "").lstrip("/\\")
    relative_url = relative_url.replace("\\", "/")
    download_url = request.build_absolute_uri(
        settings.MEDIA_URL.rstrip("/") + "/" + relative_url
    )

    return Response(
        {
            "status": True,
            "message": "Post logs and Pre-audit file uploaded successfully",
            "download_url": download_url,
        },
        status=status.HTTP_200_OK,
    )
