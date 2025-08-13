import sys
import os
import win32com.client as win32
import pythoncom
from typing import Dict, Optional
import time

class SolidWorksMetadataReader:
    def __init__(self):
        self.sw_app = None
        self.sw_model = None
    
    def connect_to_solidworks(self) -> bool:
        """Connect to SolidWorks application"""
        try:
            # Try to connect to existing SolidWorks instance first
            try:
                self.sw_app = win32.GetActiveObject("SldWorks.Application")
                print("Connected to existing SolidWorks instance")
            except:
                # Start new SolidWorks instance
                self.sw_app = win32.Dispatch("SldWorks.Application")
                print("Started new SolidWorks instance")
                time.sleep(2)  # Give SolidWorks time to start
            
            # Run SolidWorks in background (invisible)
            self.sw_app.Visible = False
            return True
            
        except Exception as e:
            print(f"Failed to connect to SolidWorks: {e}")
            return False
    
    def read_metadata(self, file_path: str) -> Dict[str, str]:
        """Read metadata from SolidWorks file"""
        metadata = {}
        
        if not self.connect_to_solidworks():
            return metadata
        
        try:
            # Normalize the file path
            file_path = os.path.abspath(file_path)
            print(f"Absolute path: {file_path}")
            
            # Determine document type
            file_ext = os.path.splitext(file_path)[1].lower()
            if file_ext == '.sldprt':
                doc_type = 1  # swDocPART
            elif file_ext == '.sldasm':
                doc_type = 2  # swDocASSEMBLY
            elif file_ext == '.slddrw':
                doc_type = 3  # swDocDRAWING
            else:
                raise ValueError(f"Unsupported file type: {file_ext}")
            
            print(f"Document type: {doc_type}")
            
            # Try simpler OpenDoc method first
            try:
                self.sw_model = self.sw_app.OpenDoc(file_path, doc_type)
            except:
                # Fallback to OpenDoc6 with minimal parameters
                self.sw_model = self.sw_app.OpenDoc6(
                    file_path,
                    doc_type,
                    1,  # swOpenDocOptions_LoadModel
                    "",
                    0,
                    0
                )
            
            if not self.sw_model:
                raise Exception("Failed to open file")
            
            print(f"Successfully opened: {os.path.basename(file_path)}")
            
            # Get custom properties
            try:
                prop_manager = self.sw_model.Extension.CustomPropertyManager("")
                if prop_manager:
                    # GetNames might return a tuple, so handle that
                    prop_names_result = prop_manager.GetNames()
                    
                    # Handle different return types
                    if prop_names_result:
                        if isinstance(prop_names_result, tuple):
                            prop_names = prop_names_result[0] if len(prop_names_result) > 0 else None
                        else:
                            prop_names = prop_names_result
                        
                        if prop_names:
                            for prop_name in prop_names:
                                try:
                                    # Try different approaches to get the property value
                                    val_out = prop_manager.Get(prop_name)
                                    if val_out:
                                        if isinstance(val_out, tuple) and len(val_out) > 0:
                                            metadata[f"Custom_{prop_name}"] = str(val_out[0]) if val_out[0] else ""
                                        else:
                                            metadata[f"Custom_{prop_name}"] = str(val_out)
                                except Exception as e:
                                    print(f"Error getting custom property {prop_name}: {e}")
                                    continue
            except Exception as e:
                print(f"Error accessing custom properties: {e}")
            
            # Get summary information
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
                    value = self.sw_model.SummaryInfo(field_id)
                    if value:
                        metadata[f"Summary_{field_name}"] = str(value)
                except:
                    continue
            
            # Get file properties
            try:
                title_result = self.sw_model.GetTitle
                title = title_result() if callable(title_result) else title_result
                if title:
                    if isinstance(title, (list, tuple)):
                        metadata["FileName"] = str(title[0]) if len(title) > 0 else ""
                    else:
                        metadata["FileName"] = str(title)
                
                path_result = self.sw_model.GetPathName
                path = path_result() if callable(path_result) else path_result
                if path:
                    if isinstance(path, (list, tuple)):
                        metadata["FilePath"] = str(path[0]) if len(path) > 0 else ""
                    else:
                        metadata["FilePath"] = str(path)
            except Exception as e:
                print(f"Error getting file properties: {e}")
            
            # Get configuration info
            try:
                config_manager = self.sw_model.ConfigurationManager
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
                        
                        config_names_result = self.sw_model.GetConfigurationNames
                        config_names = config_names_result() if callable(config_names_result) else config_names_result
                        if config_names and hasattr(config_names, '__len__'):
                            metadata["ConfigurationCount"] = str(len(config_names))
            except Exception as e:
                print(f"Error getting configuration info: {e}")
            
            # Get material info (for parts)
            if doc_type == 1:  # Part file
                try:
                    material_property = self.sw_model.MaterialPropertyValues
                    if material_property and len(material_property) > 0:
                        density = material_property[0]
                        metadata["MaterialDensity"] = str(density) if density else ""
                except Exception as e:
                    print(f"Error getting material properties: {e}")
            
            # Get mass properties (if available)
            try:
                mass_props = self.sw_model.Extension.CreateMassProperty()
                if mass_props:
                    mass = getattr(mass_props, 'Mass', None)
                    volume = getattr(mass_props, 'Volume', None)
                    surface_area = getattr(mass_props, 'SurfaceArea', None)
                    
                    if mass is not None:
                        metadata["Mass"] = f"{mass:.6f}"
                    if volume is not None:
                        metadata["Volume"] = f"{volume:.6f}"
                    if surface_area is not None:
                        metadata["SurfaceArea"] = f"{surface_area:.6f}"
            except Exception as e:
                print(f"Error getting mass properties: {e}")
                
        except Exception as e:
            print(f"Error reading metadata: {e}")
            
        finally:
            self.cleanup()
        
        return metadata
    
    def cleanup(self):
        """Clean up COM objects"""
        try:
            if self.sw_model:
                # Get the document title/name
                doc_title = self.sw_model.GetTitle()
                if isinstance(doc_title, tuple):
                    doc_title = doc_title[0] if doc_title else ""
                
                # Close the document
                self.sw_app.CloseDoc(doc_title)
                self.sw_model = None
            
            if self.sw_app:
                # Don't quit SolidWorks if it was already running
                # self.sw_app.ExitApp()
                self.sw_app = None
                
        except Exception as e:
            print(f"Cleanup warning: {e}")

def main():
    if len(sys.argv) != 2:
        print("Usage: python solidworks_metadata.py <path_to_solidworks_file>")
        print("Supported files: .sldprt, .sldasm, .slddrw")
        return
    
    file_path = sys.argv[1]
    
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return
    
    # Check file extension
    valid_extensions = ['.sldprt', '.sldasm', '.slddrw']
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext not in valid_extensions:
        print(f"Unsupported file type: {file_ext}")
        print(f"Supported types: {', '.join(valid_extensions)}")
        return
    
    print(f"Reading metadata from: {file_path}")
    print("-" * 50)
    
    reader = SolidWorksMetadataReader()
    metadata = reader.read_metadata(file_path)
    
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