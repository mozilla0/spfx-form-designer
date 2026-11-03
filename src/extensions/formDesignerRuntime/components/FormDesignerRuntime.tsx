import * as React from 'react';
import { MessageBar, MessageBarType, DefaultButton, Stack } from '@fluentui/react';
import { FormMode, IFormDefinition } from '../../../common/model';
import { FieldValues, ICurrentUser, ISpField } from '../../../common/spTypes';
import { SpService } from '../../../common/spService';
import { FormHost } from '../../../common/components/FormHost';
import { t } from '../../../common/i18n';

export interface IFormDesignerRuntimeProps {
  definition: IFormDefinition | undefined;
  fields: ISpField[];
  values: FieldValues;
  mode: FormMode;
  loadError: string;
  currentUser: ICurrentUser;
  sp: SpService;
  onSubmit: (values: FieldValues) => Promise<void>;
  onClose: () => void;
  onRequestEdit?: () => void;
}

const FormDesignerRuntime: React.FC<IFormDesignerRuntimeProps> = (props) => {
  if (props.loadError || !props.definition) {
    return (
      <Stack tokens={{ childrenGap: 12 }} styles={{ root: { padding: 24 } }}>
        <MessageBar messageBarType={MessageBarType.error}>
          {t('Das Formular konnte nicht geladen werden.')} {props.loadError}
        </MessageBar>
        <Stack>
          <DefaultButton text={t('Schliessen')} onClick={props.onClose} styles={{ root: { alignSelf: 'flex-start' } }} />
        </Stack>
      </Stack>
    );
  }
  return (
    <FormHost
      key={props.mode}
      definition={props.definition}
      fields={props.fields}
      mode={props.mode}
      initialValues={props.values}
      currentUser={props.currentUser}
      sp={props.sp}
      onSubmit={props.onSubmit}
      onCancel={props.onClose}
      onRequestEdit={props.onRequestEdit}
    />
  );
};

export default FormDesignerRuntime;
