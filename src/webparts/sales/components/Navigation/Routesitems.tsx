import * as React from "react";
import { Route, Routes, useParams } from "react-router-dom";
import { Suspense } from "react";
import Home from "../Home/Home.component";

import AssetRegister from "../Forms/AssetRegister";
import AssetIssueReturn from "../Forms/AssetReturn";


export interface RoutesProps {
    isAuthorized: boolean,
    spContext: any;
}

export interface RoutesState {
}

class RoutesItems extends React.Component<RoutesProps, RoutesState> {

    public render() {
        const WrapperHome = (props: any, component: Comment) => {
            let params = useParams();
            return <Home {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        // Masters
      
      
    
          const WrapperAssetRegister = (props: any) => {
            let params = useParams();
            return <AssetRegister {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
          const WrapperAssetReturn = (props: any) => {
            let params = useParams();
            return <AssetIssueReturn {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
         

        return (
            <Suspense fallback={<div></div>}>
                <Routes>
                    <Route path="/" element={<WrapperHome />} />
                    <Route path="/Home" element={<WrapperHome />} />
                    {/* Masters */}

                    {/* Forms */}
                    {/* <Route path="/OpportunityPortraitForm/:id?" element={<WrapperOpportunityPortrait />} /> */}
                    <Route path="/AssetRegisterForm" element={<WrapperAssetRegister/>}></Route>
                    <Route path="/AssetIssueReturnForm" element={<WrapperAssetReturn/>}></Route>

                    {/* Views */}

                </Routes>
            </Suspense>
        )
    }

}

export default RoutesItems;