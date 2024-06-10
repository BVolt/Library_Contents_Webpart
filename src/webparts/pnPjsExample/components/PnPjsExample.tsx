import * as React from 'react';
import styles from './PnPjsExample.module.scss';
import { IPnPjsExampleProps } from './IPnPjsExampleProps';
import { IFile, IResponseItem } from "./interfaces";
import { getSP } from "../pnpjsConfig";
import { SPFI} from "@pnp/sp";
import { Logger, LogLevel } from "@pnp/logging";
import { IItemUpdateResult } from "@pnp/sp/items";
import { Label, PrimaryButton, IconButton } from '@microsoft/office-ui-fabric-react-bundle';
import { Dialog, DialogType, DialogFooter, DefaultButton, TextField } from '@fluentui/react';
import { IFileAddResult} from "@pnp/sp/files";
import "@pnp/sp/attachments";
import "@pnp/sp/folders";
import "@pnp/sp/files";

export interface IAsyncAwaitPnPJsProps {
  description: string;
}

export interface IIPnPjsExampleState { 
  items: IFile[];
  errors: string[];
  newItemFile: File | undefined;
  isDeleteDialogOpen:boolean;
  isUpdateDialogOpen: boolean;
  currentItem: IFile | undefined;
  newTitle: string;
}

export default class PnPjsExample extends React.Component<IPnPjsExampleProps, IIPnPjsExampleState> {
  private LOG_SOURCE = "🅿PnPjsExample";
  private LIBRARY_NAME = "Documents";
  private _sp: SPFI;

  constructor(props: IPnPjsExampleProps) {
    super(props);
    this.state = {
      items: [],
      errors: [],
      newItemFile: undefined,
      isDeleteDialogOpen: false,
      isUpdateDialogOpen: false,
      currentItem: undefined,
      newTitle: ""
    };
    this._sp = getSP();
  }

  // Load files when the component mounts
  public async componentDidMount(): Promise<void> {
    await this._readAllFilesSize();
  }

  public render(): React.ReactElement<IAsyncAwaitPnPJsProps> {
    const totalDocs: number = this.state.items.length > 0
      ? this.state.items.reduce<number>((acc: number, item: IFile) => {
        return (acc + Number(item.Size));
      }, 0)
      : 0;
  
    return (
      <div className={styles.pnPjsExample}>
        <Label>{`${this.LIBRARY_NAME} Library Contents`}</Label>
        <PrimaryButton onClick={this._triggerFileInput} className={styles.buttonSpacing}>Upload File</PrimaryButton>
        <PrimaryButton onClick={this._updateTitles}>Update Item Titles</PrimaryButton>
        <Label>List of documents:</Label>
        <table width="100%">
          <thead>
            <tr>
              <td><strong>Title</strong></td>
              <td><strong>Name</strong></td>
              <td><strong>Size (KB)</strong></td>
              <td><strong>Actions</strong></td>
            </tr>
          </thead>
          <tbody>
            {this.state.items.map((item, idx) => (
              <tr key={idx}>
                <td>{item.Title}</td>
                <td>{item.Name}</td>
                <td>{(item.Size / 1024).toFixed(2)}</td>
                <td>
                  <IconButton 
                    iconProps={{ iconName: 'Edit' }} 
                    title="Update" 
                    ariaLabel="Update" 
                    onClick={() => this._openUpdateDialog(item)} 
                  />
                  <IconButton 
                    iconProps={{ iconName: 'Delete' }} 
                    title="Delete" 
                    ariaLabel="Delete" 
                    onClick={() => this._openDeleteDialog(item)} 
                  />
                </td>
              </tr>
            ))}
            <tr>
              {/* <td></td> */}
              <td><strong>Total:</strong></td>
              <td><strong>{(totalDocs / 1024).toFixed(2)}</strong></td>
            </tr>
          </tbody>
        </table>
  
        {/* Hidden file input for uploading files */}
        <input 
          type="file" 
          style={{ display: 'none' }} 
          ref={input => {
            this.fileInput = input
            return input
          }} 
          onChange={this._handleFileChange} 
        />

        {/* Update Title Dialog */}
      {this.state.isUpdateDialogOpen && (
        <Dialog
          hidden={!this.state.isUpdateDialogOpen}
          onDismiss={this._closeUpdateDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Update Title',
            subText: 'Please enter the new title:'
          }}
        >
          <TextField 
            value={this.state.newTitle}
            onChange={(e, newValue) => this.setState({ newTitle: newValue || "" })}
          />
          <DialogFooter>
            <PrimaryButton onClick={this._updateItemTitle} text="Update" />
            <DefaultButton onClick={this._closeUpdateDialog} text="Cancel" />
          </DialogFooter>
        </Dialog>
      )}

       {/* Delete Title Dialog */}
      {this.state.isDeleteDialogOpen && (
        <Dialog
          hidden={!this.state.isDeleteDialogOpen}
          onDismiss={this._closeDeleteDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Delete File',
            subText: 'Are you sure you want to delete this file?'
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this._deleteItem} text="Delete" />
            <DefaultButton onClick={this._closeDeleteDialog} text="Cancel" />
          </DialogFooter>
        </Dialog>
      )}
      </div>
    );
  }
  
  // Opens the update dialog and sets the current item
  private _openUpdateDialog = (item: IFile): void => {
    this.setState({ isUpdateDialogOpen: true, currentItem: item, newTitle: item.Title });
  }
  
  // Closes the update dialog and clears the current item
  private _closeUpdateDialog = (): void => {
    this.setState({ isUpdateDialogOpen: false, currentItem: undefined, newTitle: "" });
  }
  
  // Updates the title of the current item
  private _updateItemTitle = async (): Promise<void> => {
    if (!this.state.currentItem) return;
  
    try {
      await this._sp.web.lists.getByTitle(this.LIBRARY_NAME).items.getById(this.state.currentItem.Id).update({
        Title: this.state.newTitle
      });
  
      await this._readAllFilesSize();
      this._closeUpdateDialog();
    } catch (err) {
      Logger.write(`${this.LOG_SOURCE} (_updateItemTitle) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  }

  // Opens the delete dialog and sets the current item
  private _openDeleteDialog = (item: IFile): void => {
    this.setState({ isDeleteDialogOpen: true, currentItem: item});
  }

  // Closes the delete dialog and clears the current item
  private _closeDeleteDialog = (): void => {
    this.setState({ isDeleteDialogOpen: false, currentItem: undefined});
  }

  // Deletes the current item from the SharePoint library
  private _deleteItem = async (): Promise<void> => {
    if (!this.state.currentItem) return;

    try {
      await this._sp.web.lists.getByTitle(this.LIBRARY_NAME).items.getById(this.state.currentItem.Id).delete();
      await this._readAllFilesSize();
      this._closeDeleteDialog();
    } catch (err) {
      Logger.write(`${this.LOG_SOURCE} (uploadFile) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  }

  // Reference to the hidden file input element
  private fileInput: HTMLInputElement | null = null;

  // Triggers the hidden file input click
  private _triggerFileInput = (): void => {
    if (this.fileInput) {
      this.fileInput.click();
    }
  }
  
  // Handles the file input change event
  private _handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files?.[0] || null;
    if (file) {
      this.setState({ newItemFile: file });
      await this.uploadFile(file);
    }
  }
  
  // Uploads a file to the SharePoint library
  private uploadFile = async (file: File): Promise<void> => {
    try {
      const fileContent = await file.arrayBuffer();
      const fileBlob = new Blob([fileContent], { type: file.type });

      const folder = this._sp.web.lists.getByTitle(this.LIBRARY_NAME).rootFolder;

      let result: IFileAddResult;

      // Upload in chunks if file is larger than 10 MB
      if (file.size > 10485760) {
        result = await folder.files.addChunked(file.name, file, data => {
          console.log(`progress:`);
        }, true);
      } else {
        result = await folder.files.addUsingPath(file.name, fileBlob, { Overwrite: true });
      }

      console.log("File uploaded successfully:", result);

      await this._readAllFilesSize();
      this.setState({ newItemFile: undefined });
    } catch (err) {
      Logger.write(`${this.LOG_SOURCE} (_handleFileChange) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  }

  // Reads all files from the SharePoint library and updates the state
  private _readAllFilesSize = async (): Promise<void> => {
    try {  
      const response: IResponseItem[] = await this._sp.web.lists
        .getByTitle(this.LIBRARY_NAME)
        .items
        .select("Id", "Title", "FileLeafRef", "File/Length")
        .expand("File/Length")();

      
      // use map to convert IResponseItem[] into our internal object IFile[]
      const items: IFile[] = response.map((item: IResponseItem) => {
        return {
          Id: item.Id,
          Title: item.Title || "Unknown",
          Size: item.File?.Length || 0,
          Name: item.FileLeafRef
        };
      });

      // Add the items to the state
      this.setState({ items });
    } catch (err) {
      Logger.write(`${this.LOG_SOURCE} (_readAllFilesSize) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  }

  // Updates the titles of all items in the SharePoint library
  private _updateTitles = async (): Promise<void> => {
    try {
      const [batchedSP, execute] = this._sp.batched();

      const items = JSON.parse(JSON.stringify(this.state.items));

      const res: IItemUpdateResult[] = [];

      for (let i = 0; i < items.length; i++) {
        batchedSP.web.lists
          .getByTitle(this.LIBRARY_NAME)
          .items
          .getById(items[i].Id)
          .update({ Title: `${items[i].Name}-Updated` })
          .then(r => res.push(r))
          .catch(error=> {throw error});
      }
      // Executes the batched calls
      await execute();

      // Results for all batched calls are available
      for (let i = 0; i < res.length; i++) {
        const item = await res[i].item.select("Id, Title")<{ Id: number, Title: string }>();
        const stateItem = items.find((it: IFile) => it.Id === item.Id);
        if (stateItem) {
          stateItem.Title = item.Title;
        }
      }

      //Update the state which rerenders the component
      this.setState({ items });
    } catch (err) {
      Logger.write(`${this.LOG_SOURCE} (_updateTitles) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  }
}
