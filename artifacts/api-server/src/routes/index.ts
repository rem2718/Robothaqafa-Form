import { Router, type IRouter } from "express";
import healthRouter from "./health";
import formsRouter from "./forms";
import adminRouter from "./admin";

const router: IRouter = Router();

router.use(healthRouter);
router.use(formsRouter);
router.use(adminRouter);

export default router;
