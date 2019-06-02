import * as React from 'react';
import styles from './SpfxPowerAppsForm.module.scss';
import { ISpfxPowerAppsFormProps } from './ISpfxPowerAppsFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import Iframe from 'react-iframe';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { css, classNamesFunction, DefaultButton, IButtonProps, IStyle, Label, PrimaryButton } from 'office-ui-fabric-react';

export interface ISpfxPowerAppsFormState {
  showModalNew: boolean;
}


const CloseModal = () => (
  <Icon iconName="ChromeClose" className="ms-IconExample" style={{fontWeight : "bold"}}    />
);

export default class SpfxPowerAppsForm extends React.Component<ISpfxPowerAppsFormProps, ISpfxPowerAppsFormState> {

  constructor(props: any) {
    super(props);
    this.state = {
      showModalNew: false,
      };
  }

  public render(): React.ReactElement<ISpfxPowerAppsFormProps> {
    return (
      <div>
       <DefaultButton id="requestButton" onClick={this._showModalNew} text="Raise Leave Request"></DefaultButton>&nbsp;
       <Modal
          titleAriaId="titleId"
          subtitleAriaId="subtitleId"
          isOpen={this.state.showModalNew}
          onDismiss={this._closeModalNew}
          isBlocking={false}
         containerClassName="ms-modalExample-container">
          <div >
            <span ><b>Leave Request Form  </b> </span> 
            <DefaultButton onClick={this._closeModalNew} className={styles.CloseButton} ><CloseModal/></DefaultButton>
          </div>
          <div id="subtitleId" className="ms-modal-body">
            <Iframe url={"https://web.powerapps.com/webplayer/iframeapp?source=iframe&screenColor=rgba(104,101,171,1)&appId=/providers/Microsoft.PowerApps/apps/f33f9511-5001-467f-8238-fddc36665299"}
                width="1024px" height="550px"
                position="relative"
                allowFullScreen>
            </Iframe>
          </div>
        </Modal>
       </div>
    );
  }
  private _showModalNew = (): void => {
    this.setState({ showModalNew: true });
    
  };
  private _closeModalNew = (): void => {
    this.setState({ showModalNew: false });
  };
}
