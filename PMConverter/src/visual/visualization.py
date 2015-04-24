__author__ = 'Eveline'
from visual.actualcost import ActualCost
from visual.actualduration import ActualDuration
from visual.baselineshedule import BaselineSchedule
from visual.budget import CV
from visual.cost_value_metrics import CostValueMetrics
from visual.cpi import CPI
from visual.performance import Performance
from visual.resourcedistribution import ResourceDistribution
from visual.riskanalysis import RiskAnalysis
from visual.spi_t import SpiT
from visual.sv_t import SvT
from visual.spi_t_p_factor import SpiTvsPfactor


class Visualization(object):

    all = [BaselineSchedule(), "--",
           ResourceDistribution(), "--",
           RiskAnalysis(), "--",
           ActualDuration(), ActualCost(), "--",
           CostValueMetrics(), Performance(), SpiTvsPfactor(), SvT(), CV(), CPI(), SpiT()]

    def draw(self, workbook):
        raise NotImplementedError("This method is not implemented!")