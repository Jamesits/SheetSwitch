import { useEffect, useRef } from "react";
import usePromise from "react-use-promise";

type BindEventCallback<T> = (context: Excel.RequestContext) => OfficeExtension.EventHandlers<T>
type EventCallback<T> = (args: T) => Promise<any>

export const useExcelEvent = <T>(bindEvent: BindEventCallback<T>, callback: EventCallback<T>, deps?: React.DependencyList) => {
    useEffect(() => {
        console.log("Binded")
        let ref: OfficeExtension.EventHandlerResult<T>;
        Excel.run((context) => {
            console.log("run")
            const handler = bindEvent(context)
            ref = handler.add(callback);
            return context.sync();
        })
        return () => {
            console.log("Removed")
            if (ref) {
                ref.remove();
            }
        }
    }, deps)
}
