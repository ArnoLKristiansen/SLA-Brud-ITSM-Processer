import { PropsWithChildren, useEffect } from "react";
import { PowerApps } from "@microsoft/power-apps";

export function PowerProvider({ children }: PropsWithChildren) {
    useEffect(() => {
        // Initialize Power Platform runtime
        PowerApps.initialize();
    }, []);

    return <>{children}</>;
}