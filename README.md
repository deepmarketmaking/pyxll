# Installation
1. Install PyXLL from [PyXLL Official Site](https://www.pyxll.com/index.html)  
2. Download the latest version of code from github:
    1. click on **code** button
    2. click on **Download ZIP**
    ![](./images/github-download.png?raw=true "")
3. Extract the zip.
4. Update the `pyxll.cfg` file in PyXLL installation folder. (You can check pyxll.cfg from zip file for reference)
    - Find the `executable` setting and update it to point to your Python installation.
        ![](./images/config-executable2.png?raw=true "")
    - Find `pythonpath` and add path to source folder extracted from zip file. For example `C:\Users\<your user name>\deepmm\pyxll\source`
      ![](./images/cfg-pythonpath.png?raw=true "")
    - Find `modules` and add 
        ```
        main
        websocket_handler
        ui.connection_status_ribbon
        ui.connection_status_ribbon_config
        websocket_event_listener
        subscription_manager
        ```
        each on new line
        ![](./images/cfg-modules.png?raw=true "")
    - Find `ribbon` and add  path to ribbon.xml from source folder `C:\Users\<your user name>\deepmm\pyxll\source\ui\ribbon.xml`
      ![](./images/config-ribbon2.png?raw=true "")
5. Install dependencies from `requirements.txt` by runing `pip install -r requirements.txt` command in terminal


# Working in excel
1. Open Excel and navigate to the **DeepMM** tab.
![](./images/deepmm-in-menu.png?raw=true "")
2. Click the **DeepMM** tab.
3. Click the **Configuration** button
![](./images/configure-button.png?raw=true "")
4. **Map columns** in the configuration popup:
    - Select the identifier type (CUSIP, FIGI, ISIN).
    - Match it with the corresponding Excel column (e.g., provide the column letter).
    - Repeat for all required fields.
![](./images/configure-popup.png?raw=true "")
5. Save the configuration.
6. Click the **Login** button.
![](./images/login-button.png?raw=true "")
7. Enter your **email and password**, then click **Login**.
![](./images/login-popup.png?raw=true "")
8. After logging in, data synchronization with DeepMM will start automatically.
    - **Note**: You must log in each time you open Excel.
![](./images/results.png.png?raw=true "")
9. **Adding more worksheets**:
    - Add a new worksheet and configure it to start syncing data.
![](./images/configure-button.png?raw=true "")
10. To stop syncing data, click the **Clear Configuration** button.
![](./images/clear-configuration.png.png?raw=true "")
11. **Check WebSocket connection status:**
    - The **Connected** status indicates whether the WebSocket connection is active.
    - This status is **not** related to your login credentials.


# Expected data format
 - **Identifier** - **FIGI/CUSIP/ISIN**
 - **Side**:
    - `bid`
    - `offer`
    - `dealer`
 - **Quantity**:
    - `1 000`
    - `10 000`
    - `100 000`
    - `250 000`
    - `500 000`
    - `1 000 000`
    - `2 000 000`
    - `3 000 000`
    - `4 000 000`
    - `5 000 000`
  - **Label**:
    - `price`
    - `ytm`
    - `spread`
  - **ATS**:
    - `Y`
    - `N`

Invalid rows will be skipped


# DeepMM api examples
https://github.com/deepmarketmaking/api
