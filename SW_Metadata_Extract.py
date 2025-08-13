import sys
import os
import win32com.client as win32
from typing import Dict
import time


def connect_to_solidworks():
    """Connect to SolidWorks application"""
    try:
        # Try to connect to existing SolidWorks instance first
        try:
            sw_app = win32.GetActiveObject("SldWorks.Application")
            print("Connected to existing SolidWorks instance")
        except:
            # Start new SolidWorks instance
            sw_app = win32.Dispatch("SldWorks.Application")
            print("Started new SolidWorks instance")
            time.sleep(2)  # Give SolidWorks time to start
        
        # Run SolidWorks in background (invisible)
        sw_app.Visible = False
        return sw_app
        
    except Exception as e:
        print(f"Failed to connect to SolidWorks: {e}")
        return None


def open_solidworks_file(sw_app, file_path: str):
    """Open a SolidWorks part file"""
    try:
        # Normalize the file path
        file_path = os.path.abspath(file_path)
        print(f"Absolute path: {file_path}")
        
        # Check for part file only
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext != '.sldprt':
            raise ValueError(f"Only SolidWorks part files (.sldprt) are supported. Got: {file_ext}")
        
        doc_type = 1  # swDocPART
        print("Document type: Part file")
        
        # Use OpenDoc (simpler and works reliably)
        sw_model = sw_app.OpenDoc(file_path, doc_type)
        
        if not sw_model:
            raise Exception("Failed to open file")
        
        print(f"Successfully opened: {os.path.basename(file_path)}")
        return sw_model
        
    except Exception as e:
        print(f"Error opening file: {e}")
        raise


def extract_custom_properties(sw_model) -> Dict[str, str]:
    """Extract custom properties from SolidWorks model"""
    metadata = {}
    key_custom_props = ["Weight", "Material", "Thickness", "Description"]
    found_custom_props = set()
    
    try:
        prop_manager = sw_model.Extension.CustomPropertyManager("")
        if prop_manager:
            # Get all custom property names - handle callable vs tuple issue
            try:
                # Try as method call first
                get_names_result = prop_manager.GetNames()
            except TypeError:
                # If it's not callable, it might be returning the result directly
                get_names_result = getattr(prop_manager, 'GetNames', ())
            
            prop_names = []
            if get_names_result and isinstance(get_names_result, tuple):
                prop_names = list(get_names_result)
            
            print(f"Found {len(prop_names)} custom properties: {prop_names}")
            
            # Extract all custom properties
            for prop_name in prop_names:
                try:
                    # Use Get4 method to get both raw and evaluated values (like C# example)
                    # Get4(PropertyName, UseCache, out Value, out EvaluatedValue)
                    get_result = prop_manager.Get4(prop_name, False)
                    
                    if get_result and isinstance(get_result, tuple) and len(get_result) >= 3:
                        # get_result = (status, raw_value, evaluated_value)
                        status = get_result[0]
                        raw_value = get_result[1] if get_result[1] else ""
                        evaluated_value = get_result[2] if get_result[2] else ""
                        
                        # Prefer evaluated value over raw value
                        final_value = evaluated_value if evaluated_value else raw_value
                        metadata[f"Custom_{prop_name}"] = final_value
                        
                        print(f"  {prop_name}: '{raw_value}' -> '{evaluated_value}'")
                        
                        if prop_name in key_custom_props:
                            found_custom_props.add(prop_name)
                    else:
                        # Fallback to regular Get method
                        get_result = prop_manager.Get(prop_name)
                        if get_result and isinstance(get_result, str):
                            metadata[f"Custom_{prop_name}"] = get_result
                            if prop_name in key_custom_props:
                                found_custom_props.add(prop_name)
                        
                except Exception as e:
                    print(f"Error getting custom property {prop_name}: {e}")
                    # Try fallback method
                    try:
                        get_result = prop_manager.Get(prop_name)
                        if get_result and isinstance(get_result, str):
                            metadata[f"Custom_{prop_name}"] = get_result
                    except:
                        pass
            
            # Ensure key custom properties are always present
            for key_prop in key_custom_props:
                if key_prop not in found_custom_props:
                    metadata[f"Custom_{key_prop}"] = ""
                    
    except Exception as e:
        print(f"Error accessing custom properties: {e}")
        # Ensure key properties are present even if there's an error
        for key_prop in key_custom_props:
            metadata[f"Custom_{key_prop}"] = ""
    
    return metadata


def extract_summary_info(sw_model) -> Dict[str, str]:
    """Extract summary information from SolidWorks model"""
    metadata = {}
    summary_info_fields = {
        0: "Title",
        1: "Subject", 
        2: "Author",
        3: "Keywords",
        4: "Comments",
        5: "LastSavedBy",
        6: "RevisionNumber",
        9: "CreatedDate",
        10: "ModifiedDate",
        11: "LastPrintedDate"
    }
    
    for field_id, field_name in summary_info_fields.items():
        try:
            value = sw_model.SummaryInfo(field_id)
            if value:
                metadata[f"Summary_{field_name}"] = str(value)
        except:
            continue
    
    return metadata


def extract_file_properties(sw_model) -> Dict[str, str]:
    """Extract file properties from SolidWorks model"""
    metadata = {}
    
    try:
        title_result = sw_model.GetTitle
        title = title_result() if callable(title_result) else title_result
        if title:
            if isinstance(title, (list, tuple)):
                metadata["FileName"] = str(title[0]) if len(title) > 0 else ""
            else:
                metadata["FileName"] = str(title)
        
        path_result = sw_model.GetPathName
        path = path_result() if callable(path_result) else path_result
        if path:
            if isinstance(path, (list, tuple)):
                metadata["FilePath"] = str(path[0]) if len(path) > 0 else ""
            else:
                metadata["FilePath"] = str(path)
    except Exception as e:
        print(f"Error getting file properties: {e}")
    
    return metadata


def extract_configuration_info(sw_model) -> Dict[str, str]:
    """Extract configuration information from SolidWorks model"""
    metadata = {}
    
    try:
        config_manager = sw_model.ConfigurationManager
        if config_manager:
            active_config = config_manager.ActiveConfiguration
            if active_config:
                config_name_result = active_config.Name
                config_name = config_name_result() if callable(config_name_result) else config_name_result
                if config_name:
                    if isinstance(config_name, (list, tuple)):
                        metadata["ActiveConfiguration"] = str(config_name[0]) if len(config_name) > 0 else ""
                    else:
                        metadata["ActiveConfiguration"] = str(config_name)
    except Exception as e:
        print(f"Error getting configuration info: {e}")
    
    return metadata


def extract_material_info(sw_model) -> Dict[str, str]:
    """Extract material information from SolidWorks model"""
    metadata = {}
    
    try:
        material_property = sw_model.MaterialPropertyValues
        if material_property and len(material_property) > 0:
            density = material_property[0]
            metadata["MaterialDensity"] = str(density) if density else ""
    except Exception as e:
        print(f"Error getting material properties: {e}")
    
    return metadata


def read_metadata(file_path: str) -> Dict[str, str]:
    """Read metadata from SolidWorks file, always including key custom properties if present, and all native metadata."""
    metadata = {}
    sw_app = None
    sw_model = None
    
    try:
        # Connect to SolidWorks
        sw_app = connect_to_solidworks()
        if not sw_app:
            return metadata
        
        # Open the file
        sw_model = open_solidworks_file(sw_app, file_path)
        
        # Extract all metadata
        metadata.update(extract_custom_properties(sw_model))
        metadata.update(extract_summary_info(sw_model))
        metadata.update(extract_file_properties(sw_model))
        metadata.update(extract_configuration_info(sw_model))
        metadata.update(extract_material_info(sw_model))
        
    except Exception as e:
        print(f"Error reading metadata: {e}")
    finally:
        # Clean up COM objects
        if sw_model:
            sw_model = None
        if sw_app:
            sw_app = None
    
    return metadata


def main():
    if len(sys.argv) != 2:
        print("Usage: python SW_Metadata_Extract.py <path_to_solidworks_part_file>")
        print("Supported files: .sldprt only")
        return
    
    file_path = sys.argv[1]
    
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return
    
    # Check file extension
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext != '.sldprt':
        print(f"Unsupported file type: {file_ext}")
        print("This script only supports SolidWorks part files (.sldprt)")
        return
    
    print(f"Reading metadata from: {file_path}")
    print("-" * 50)
    
    metadata = read_metadata(file_path)
    
    if metadata:
        print("SolidWorks File Metadata:")
        print("=" * 50)
        for key, value in sorted(metadata.items()):
            if value:  # Only show non-empty values
                print(f"{key:<25}: {value}")
        
        print(f"\nTotal metadata fields found: {len([v for v in metadata.values() if v])}")
    else:
        print("No metadata could be extracted.")


if __name__ == "__main__":
    main()
