import { FormEngine, IControlState } from '../../../form/FormEngine';
import { IFormValues } from '../../../form/formValues';
import { IFieldInfo } from '../../../model/IFieldInfo';
import { IControl } from '../../../model/IFormDefinition';
import { ListService } from '../../../services/ListService';

export interface IFieldControlProps {
  id: string;
  control: IControl;
  field: IFieldInfo;
  value: unknown;
  values: IFormValues;
  state: IControlState;
  label: string;
  engine: FormEngine;
  /** Undefined in an offline preview: pickers then show no suggestions. */
  listService?: ListService;
  onChange: (value: unknown) => void;
}
