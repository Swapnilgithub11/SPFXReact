import * as React from 'react';
import styles from './TrainingModule.module.scss';
import { ITrainingModuleProps } from './ITrainingModuleProps';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
//import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { PeoplePicker , PrincipalType} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ITrainingModule } from "../Model/ITrainingModule";
import pnp, { List, ListEnsureResult, ItemAddResult, FieldAddResult } from "sp-pnp-js";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/components/Button';
//import { FieldUserRenderer } from "@pnp/spfx-controls-react/lib/FieldUserRenderer";
//import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
//import { FieldRendererHelper } from '@pnp/spfx-controls-react/lib/Utilities';
//import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DatePicker , DayOfWeek, IDatePickerStrings  } from 'office-ui-fabric-react/lib/DatePicker';
import { addMonths, addYears } from 'office-ui-fabric-react/lib/utilities/dateMath/DateMath';
import * as jquery from 'jquery';
import {
  assign,
  autobind
} from 'office-ui-fabric-react/lib/Utilities';
export default class TrainingModule extends React.Component<ITrainingModuleProps, ITrainingModule> {
  
  constructor(props) {
    super(props);
    this.handleTitle = this.handleTitle.bind(this);
    this.handleDesc = this.handleDesc.bind(this);
    this.createItem = this.createItem.bind(this);
    this._getTrainer = this._getTrainer.bind(this);
    this.state = {
      name: "",
      description: "",
      selectedItems: [],
      hideDialog: true,
      showPanel: false,
      dpselectedItem: undefined,
      dpselectedItems: [],
      disableToggle: false,
      defaultChecked: false,
      termKey: undefined,
      userIDs: [],
      TrainerId: [],
      pplPickerType: "",
      status: "",
      isChecked: false,
      required: "This is required",
      onSubmission: false,
      firstDayOfWeek: DayOfWeek.Sunday,
      Id: undefined,
      Title:'',
      Category:'',
      value: null,
      items: [ 
        { 
          "Training Id": "",
          "Training Title": "", 
          "Trainer": [
            {
              "TrainerTitle" : "",
            }
          ], 
          "TrainingDate" : "", 
          "Category":"",
          "Status":"",
        } 
      ] 
    };
  }

  private dropdownOptions: IDropdownOption[];
  private listsFetched: boolean;
  
  public render(): React.ReactElement<ITrainingModuleProps> {
    const { dpselectedItem, dpselectedItems } = this.state;
    const { name, description } = this.state;
    const { firstDayOfWeek } = this.state;
    const today: Date = new Date(Date.now());
    const minDate: Date = addMonths(today, -1);
    const maxDate: Date = addYears(today, 1);
    const DayPickerStrings: IDatePickerStrings = {
      months: [
        'January',
        'February',
        'March',
        'April',
        'May',
        'June',
        'July',
        'August',
        'September',
        'October',
        'November',
        'December'
      ],
      shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

      days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

      shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

      goToToday: 'Go to today',
      prevMonthAriaLabel: 'Go to previous month',
      nextMonthAriaLabel: 'Go to next month',
      prevYearAriaLabel: 'Go to previous year',
      nextYearAriaLabel: 'Go to next year',
      isRequiredErrorMessage: "This is required",
    };

    pnp.setup({
      spfxContext: this.props.context
    });
    return (
      <form>
        <div className={styles.trainingModule}>
          <div className={styles.container}>
            <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-white ${styles.row}`}>
              <div className="ms-Grid-col ms-u-sm4 block">
                <label className="ms-Label">Training Title</label>
              </div>
              <div className="ms-Grid-col ms-u-sm8 block">
                <TextField value={this.state.name} required={true} onChanged={this.handleTitle}
                  errorMessage={(this.state.name.length === 0 && this.state.onSubmission === true) ? this.state.required : ""} />
              </div>
              <div className="ms-Grid-col ms-u-sm4 block">
                <label className="ms-Label">Training Description</label>
              </div>
              <div className="ms-Grid-col ms-u-sm8 block">
                <TextField multiline autoAdjustHeight value={this.state.description} onChanged={this.handleDesc}
                />
              </div>
              <div className="ms-Grid-col ms-u-sm4 block">
                <label className="ms-Label">Training Date</label>
              </div>
              <div className="ms-Grid-col ms-u-sm8 block">
                <DatePicker
                  firstDayOfWeek={firstDayOfWeek}
                  placeholder="Select a date..."
                  isRequired={true}
                  onSelectDate={this._onSelectDate}
                  minDate={minDate}
                  maxDate={maxDate}
                  strings={DayPickerStrings}
                  onAfterMenuDismiss={() => console.log('onAfterMenuDismiss called')}
                  
                />
              </div>
              <div className="ms-Grid-col ms-u-sm4 block">
                <label className="ms-Label">Category</label><br />
              </div>
              <div className="ms-Grid-col ms-u-sm8 block">
                <Dropdown 
                  required={true}
                  placeHolder="Select an Option"
                  label=""
                  id="component"
                  selectedKey={dpselectedItem ? dpselectedItem.key : undefined}
                  ariaLabel="Basic dropdown example"
                  options= {[
                    { key: 'Web Dev', text: 'Web Dev' },
                    { key: 'Desktop Dev', text: 'Desktop Dev' },
                    { key: 'Production Management', text: 'Production Management' }
                  ]}
                  onChanged={this._changeState}
                  onFocus={this._log('onFocus called')}
                  onBlur={this._log('onBlur called')}
                  errorMessage={(this.state.value === null && this.state.onSubmission === true) ? this.state.required : ""}
                />
              </div>
              <div className="ms-Grid-col ms-u-sm4 block">
                <label className="ms-Label">Trainer</label>
              </div>
              <div className="ms-Grid-col ms-u-sm8 block">
                <PeoplePicker
                  context= {this.props.context}
                  titleText=" "
                  personSelectionLimit={1}
                  groupName={''} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  isRequired={true}
                  disabled={false}
                  selectedItems={this._getTrainer}
                  principleTypes={[PrincipalType.User]}
                  errorMessage={(this.state.TrainerId.length === 0 && this.state.onSubmission === true) ? this.state.required : " "}
                  errorMessageclassName={styles.hideElementManager}
                />
              </div>
              <div className="ms-Grid-col ms-u-sm2 block">
                <PrimaryButton text="Create" onClick={() => { this.validateForm(); }} />
              </div>
              <div className="ms-Grid-col ms-u-sm2 block">
                <DefaultButton text="Cancel" onClick={() => { this.setState({}); }} />
              </div>
              <div>
                <Panel
                  isOpen={this.state.showPanel}
                  type={PanelType.smallFixedFar}
                  onDismiss={this._onClosePanel}
                  isFooterAtBottom={false}
                  headerText="Are you sure you want to create training ?"
                  closeButtonAriaLabel="Close"
                  onRenderFooterContent={this._onRenderFooterContent}
                ><span>Please check the details filled and click on Confirm button to create training.</span>
                </Panel>
              </div>
              <Dialog
                hidden={this.state.hideDialog}
                onDismiss={this._closeDialog}
                dialogContentProps={{
                  type: DialogType.largeHeader,
                  title: 'Request Submitted Successfully',
                  subText: ""
                }}
                modalProps={{
                  titleAriaId: 'myLabelId',
                  subtitleAriaId: 'mySubTextId',
                  isBlocking: false,
                  containerClassName: 'ms-dialogMainOverride'
                }}>
                <div dangerouslySetInnerHTML={{ __html: this.state.status }} />
                <DialogFooter>
                  <PrimaryButton onClick={() => this.gotoHomePage()} text="Okay" />
                </DialogFooter>
              </Dialog>
            </div>
          </div>
        </div>
        <div className={styles.trainingModule} >
          <br></br>
          <div className={styles.headerCaptionStyle} >Training Details</div>
          <div className={styles.tableStyle} >  
            <div className={styles.headerStyle} >
              <div className={styles.CellStyle}>Training ID</div>
              <div className={styles.CellStyle}>Training Title</div> 
              <div className={styles.CellStyle}>Trainer</div> 
              <div className={styles.CellStyle}>Training Date</div> 
              <div className={styles.CellStyle}>Category</div>
              <div className={styles.CellStyle}>Status</div>                     
            </div> 
              {this.state.items.map((item,key) => { 
                return (<div className={styles.rowStyle} key={key}> 
                    <div className={styles.CellStyle}>{item.Id}</div>
                    <div className={styles.CellStyle}>{item.Title}</div> 
                    <div className={styles.CellStyle}>{item.Trainer.Title}</div> 
                    <div className={styles.CellStyle}>{item.TrainingDate != null ?new Date(item.TrainingDate).toLocaleDateString():""}</div>
                    <div className={styles.CellStyle}>{item.Category}</div>
                    <div className={styles.CellStyle}>{item.Status}</div>
                  </div>); 
              })}                     
          </div> 
          </div>
      </form>
    );
  }
  
  public componentDidMount(){ 
    var reactHandler = this; 
    jquery.ajax({ 
        url: `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listname}')/items?$select=*,Trainer/Title&$expand=Trainer`, 
        type: "GET", 
        headers:{'Accept': 'application/json; odata=verbose;'}, 
        success: (resultData) => { 
          reactHandler.setState({ 
            items: resultData.d.results
          }); 
        }, 
        error : (jqXHR, textStatus, errorThrown) => { 
        } 
    }); 
  } 

  private _onSelectDate = (date: Date | null | undefined): void => {
    this.setState({ value: date });
  }

  private _getTrainer(items: any[]) {
    this.state.TrainerId.length = 0;
    for (let item in items) {
      this.state.TrainerId.push(items[item].id);
      console.log(items[item].id);
    } 
  }

  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton onClick={this.createItem} style={{ marginRight: '8px' }}>
          Confirm
      </PrimaryButton>
        <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
      </div>
    );
  }

  private _log(str: string): () => void {
    return (): void => {
      console.log(str);
    };
  }

  private _onClosePanel = () => {
    this.setState({ showPanel: false });
  }

  private _onShowPanel = () => {
    this.setState({ showPanel: true });
  }

  private _changeSharing(checked: any): void {
    this.setState({ defaultChecked: checked });
  }

  private _changeState = (item: IDropdownOption): void => {
    console.log('here is the things updating...' + item.key + ' ' + item.text + ' ' + item.selected);
    this.setState({ dpselectedItem: item });
    if (item.text == "Employee") {
      this.setState({ defaultChecked: false });
      this.setState({ disableToggle: true });
    }
    else {
      this.setState({ disableToggle: false });
    }
  }

  private handleTitle(value: string): void {
    return this.setState({
      name: value
    });
  }

  private handleDesc(value: string): void {
    return this.setState({
      description: value
    });
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

  private _showDialog = (status: string): void => {
    this.setState({ hideDialog: false });
    this.setState({ status: status });
  }

  private validateForm(): void {
    let allowCreate: number = 0;
    this.setState({ onSubmission: true });
    
    if (this.state.name.length === 0) {
      allowCreate++;
    }
    if(this.state.value === null){
      allowCreate++;
    }
    if(this.state.TrainerId.length === 0){
      allowCreate++;
    }

    if (allowCreate === 0) {
      this._onShowPanel();
    }
    else {
      //do nothing
    }
  }

 private createItem(): void {  
  pnp.setup({
    sp: {
      baseUrl: this.props.siteUrl,
    }
  });
  this._onClosePanel();
  this._showDialog("Submitting Request");
  pnp.sp.web.lists.getByTitle(this.props.listname).items.add({  
    Title: this.state.name,
    Description: this.state.description,
    Category: this.state.dpselectedItem.key,
    TrainingDate: new Date(this.state.value),
    TrainerId:Number(this.state.TrainerId[0]),
  });
  }

  private gotoHomePage(): void {
    window.location.replace(this.props.siteUrl);
  }
}