package com.inductiveautomation.apachepoi;

import com.inductiveautomation.ignition.common.licensing.LicenseState;
import com.inductiveautomation.ignition.common.script.ScriptManager;
import com.inductiveautomation.ignition.gateway.clientcomm.ClientReqSession;
import com.inductiveautomation.ignition.gateway.model.AbstractGatewayModuleHook;
import com.inductiveautomation.ignition.gateway.model.GatewayContext;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class GatewayHook extends AbstractGatewayModuleHook {

    private final Logger logger = LoggerFactory.getLogger(getClass());

    @Override
    public void setup(GatewayContext gatewayContext) {

    }

    @Override
    public void startup(LicenseState licenseState) {

    }

    @Override
    public boolean isFreeModule() {
        return true;
    }


    @Override
    public void shutdown() {

    }

    @Override
    public void initializeScriptManager(ScriptManager manager) {
        manager.addScriptModule("system.excel", ExcelFunctions.class);
    }


    @Override
    public Object getRPCHandler(ClientReqSession session, Long projectId) {
        return new RPCHandler();
    }

}
