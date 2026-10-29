import * as React from "react";
import { Route, Routes, useParams } from "react-router-dom";
import { Suspense } from "react";
import Home from "../Home/Home.component";
import Location from "../Masters/Location";
import Brand from "../Masters/Brand";
import AssetType from "../Masters/AssetType";
import Vendor from "../Masters/Vendor";
import AssetModel from "../Masters/AssetModel";

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
          const WrapperLocation = (props: any) => {
            let params = useParams();
            return <Location {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperBrand = (props: any) => {
            let params = useParams();
            return <Brand {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
          const WrapperAssetType = (props: any) => {
            let params = useParams();
            return <AssetType {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperAssetModel = (props: any) => {
            let params = useParams();
            return <AssetModel {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }
        const WrapperVendor = (props: any) => {
            let params = useParams();
            return <Vendor {...this.context} {...this.props} {...{ ...props, match: { params } }} />
        }

        return (
            <Suspense fallback={<div></div>}>
                <Routes>
                    <Route path="/" element={<WrapperHome />} />
                    <Route path="/Home" element={<WrapperHome />} />
                    {/* Masters */}
                    <Route path="/Location" element={<WrapperLocation />} />
                    <Route path="/Brand" element={<WrapperBrand />} />
                    <Route path="/AssetType" element={<WrapperAssetType />} />
                    <Route path="/Vendor" element={<WrapperVendor />} />
                    <Route path="/AssetModel" element={<WrapperAssetModel />} />
                    {/* Forms */}
                    {/* <Route path="/OpportunityPortraitForm/:id?" element={<WrapperOpportunityPortrait />} /> */}
                    
                  

                    {/* Views */}

                </Routes>
            </Suspense>
        )
    }

}

export default RoutesItems;