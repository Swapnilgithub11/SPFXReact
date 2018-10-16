
​import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import './DatePicker.Examples.scss';

export interface ITrainingModule {
    selectedItems: any[];
    name: string;
    description: string;
    dpselectedItem?: { key: string | number | undefined };
    termKey?: string | number;
    dpselectedItems: IDropdownOption[];
    disableToggle: boolean;
    defaultChecked: boolean;
    pplPickerType:string;
    userIDs: number[];
    TrainerId: number[];
    hideDialog: boolean;
    status: string;
    isChecked: boolean;
    showPanel: boolean;
    required:string;
    onSubmission:boolean;
    firstDayOfWeek?: DayOfWeek;
    Id: number;
    Title: string;
    Category: string;
    value: any;
    items: any[];
}

